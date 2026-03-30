// main.go
package main

import (
	"crypto/sha256"
	"encoding/csv"
	"encoding/json"
	"fmt"
	"math"
	"os"
	"regexp"
	"sort"
	"strconv"
	"strings"
	"time"

	"github.com/charmbracelet/bubbles/spinner"
	"github.com/charmbracelet/bubbles/textinput"
	tea "github.com/charmbracelet/bubbletea"
	"github.com/charmbracelet/lipgloss"
	"github.com/jung-kurt/gofpdf"
)

// Theme controls colors and watermark settings for PDF output.
type Theme struct {
	Primary          string
	Secondary        string
	WatermarkText    string
	WatermarkOpacity float64
	WatermarkSize    float64
	WatermarkAngle   float64
}

var theme = Theme{
	Primary:          "#0B6B3A",
	Secondary:        "#ECFDF5",
	WatermarkText:    "TAX COMPANION",
	WatermarkOpacity: 0.08,
	WatermarkSize:    48.0,
	WatermarkAngle:   45.0,
}

// hexToRGB converts a 6-digit hex color string (e.g. "#0B6B3A") to r,g,b ints.
func hexToRGB(hex string) (int, int, int) {
	hex = strings.TrimPrefix(hex, "#")
	if len(hex) != 6 {
		return 0, 0, 0
	}
	v, err := strconv.ParseInt(hex, 16, 32)
	if err != nil {
		return 0, 0, 0
	}
	r := int((v >> 16) & 0xFF)
	g := int((v >> 8) & 0xFF)
	b := int(v & 0xFF)
	return r, g, b
}

const (
	appName   = "Tax Companion (Bangladesh)"
	appAuthor = "rhshourav"
	appGitHub = "github.com/rhshourav"
)

type appState int

const (
	stateInput appState = iota
	stateLoading
	stateResult
	statePrompt
)

type promptMode string

const (
	promptNone      promptMode = ""
	promptSavePDF   promptMode = "save_pdf"
	promptSaveJSON  promptMode = "save_json"
	promptExportCSV promptMode = "export_csv"
	promptExportMD  promptMode = "export_md"
	promptLoadJSON  promptMode = "load_json"
)

type TaxSlab struct {
	Limit float64 `json:"limit"` // negative or zero = unlimited
	Rate  float64 `json:"rate"`
	Label string  `json:"label"`
}

type Config struct {
	TaxYear               string          `json:"tax_year"`
	SalaryExemptionCap    float64         `json:"salary_exemption_cap"`
	MinimumTax            float64         `json:"minimum_tax"`
	FestivalBonusRatio    float64         `json:"festival_bonus_ratio"`
	RebatePctOfTaxable    float64         `json:"rebate_pct_of_taxable"`
	RebatePctOfInvestment float64         `json:"rebate_pct_of_investment"`
	RebateMaxAmount       float64         `json:"rebate_max_amount"`
	SurchargeThreshold    float64         `json:"surcharge_threshold"`
	SurchargeAutoTiers    []SurchargeTier `json:"surcharge_auto_tiers"`
	TaxSlabs              []TaxSlab       `json:"tax_slabs"`
}

type SurchargeTier struct {
	MaxNetWealth float64 `json:"max_net_wealth"`
	Rate         float64 `json:"rate"`
}

type SessionData struct {
	Version   int      `json:"version"`
	Timestamp string   `json:"timestamp"`
	Inputs    []string `json:"inputs"`
	Step      int      `json:"step"`
	Section   int      `json:"section"`
}

type SavedFile struct {
	Hash    string      `json:"hash"`
	Session SessionData `json:"session"`
}

type TaxLine struct {
	Label  string
	Amount float64
	Rate   float64
	Tax    float64
}

type Report struct {
	GeneratedAt time.Time

	TaxYear string

	SalaryInput         float64
	BonusIncluded       bool
	CustomSalary        bool
	BaseGrossSalary     float64
	HouseRentAllowance  float64
	MedicalAllowance    float64
	ConveyanceAllowance float64
	BonusAmount         float64
	TotalComp           float64
	SalaryExempt        float64
	TaxableSalary       float64

	CurrentTaxBeforeRebate float64
	RebateEligibleInvest   float64
	RebateAmount           float64
	CurrentTaxAfterRebate  float64

	PreviousGrossInput  float64
	PreviousTax         float64
	CombinedBeforeSurch float64

	ApplySurcharge  bool
	SurchargeRate   float64
	SurchargeAmount float64
	FinalTax        float64

	TotalExpense float64
	ExpensePcts  map[string]int
	ExpenseAmts  map[string]int

	TotalAssets      float64
	TotalLiabilities float64
	NetWealthCurrent float64
	OpeningNetWealth float64
	WealthIncrease   float64
	ExpectedSavings  float64
	WealthDifference float64
	WealthStatus     string

	AuditRisk string // "LOW", "MEDIUM", "HIGH"

	ReverseEnabled        bool
	TargetNetTakeHome     float64
	Deductions            float64
	EstimatedGrossFromNet float64

	TaxLines     []TaxLine
	PrevTaxLines []TaxLine
	ReportText   string
}

type model struct {
	state appState

	step    int
	section int

	inputs []textinput.Model
	spin   spinner.Model

	width  int
	height int

	showInfo     bool
	infoOffset   int
	resultOffset int

	report Report

	config Config
	notice string

	promptActive bool
	promptMode   promptMode
	prompt       textinput.Model
}

type calcDoneMsg struct {
	report Report
}

type saveDoneMsg struct {
	message string
}

type loadDoneMsg struct {
	model   model
	message string
}

var (
	titleStyle = lipgloss.NewStyle().
			Bold(true).
			Foreground(lipgloss.Color("#FFFFFF")).
			Background(lipgloss.Color("#0B6B3A")).
			Padding(0, 1).
			MarginBottom(1)

	subTitleStyle = lipgloss.NewStyle().
			Bold(true).
			Foreground(lipgloss.Color("#0B6B3A")).
			Background(lipgloss.Color("#ECFDF5")).
			Padding(0, 1).
			MarginTop(1).
			MarginBottom(1)

	helpStyle = lipgloss.NewStyle().
			Foreground(lipgloss.Color("#666666"))

	keyStyle = lipgloss.NewStyle().
			Foreground(lipgloss.Color("#0EA5E9")).Bold(true)

	moneyStyle = lipgloss.NewStyle().Foreground(lipgloss.Color("#16A34A"))
	taxStyle   = lipgloss.NewStyle().Foreground(lipgloss.Color("#DC2626")).Bold(true)
	warnStyle  = lipgloss.NewStyle().Foreground(lipgloss.Color("#D97706")).Bold(true)
	okStyle    = lipgloss.NewStyle().Foreground(lipgloss.Color("#0F766E")).Bold(true)

	boxStyle = lipgloss.NewStyle().
			Border(lipgloss.RoundedBorder()).
			Padding(1, 1)
)

func defaultConfig() Config {
	return Config{
		TaxYear:               "2025-26",
		SalaryExemptionCap:    500000,
		MinimumTax:            3000,
		FestivalBonusRatio:    42714.0 / 344475.0, // fallback (kept for backwards compatibility)
		RebatePctOfTaxable:    0.03,
		RebatePctOfInvestment: 0.15,
		RebateMaxAmount:       1000000,
		SurchargeThreshold:    40000000,
		SurchargeAutoTiers: []SurchargeTier{
			{MaxNetWealth: 50000000, Rate: 0.10},
			{MaxNetWealth: 100000000, Rate: 0.20},
			{MaxNetWealth: math.MaxFloat64, Rate: 0.35},
		},
		TaxSlabs: []TaxSlab{
			{Limit: 350000, Rate: 0.00, Label: "First Tk 3,50,000 (0%)"},
			{Limit: 100000, Rate: 0.05, Label: "Next Tk 1,00,000 (5%)"},
			{Limit: 400000, Rate: 0.10, Label: "Next Tk 4,00,000 (10%)"},
			{Limit: 500000, Rate: 0.15, Label: "Next Tk 5,00,000 (15%)"},
			{Limit: 500000, Rate: 0.20, Label: "Next Tk 5,00,000 (20%)"},
			{Limit: 2000000, Rate: 0.25, Label: "Next Tk 20,00,000 (25%)"},
			{Limit: -1, Rate: 0.25, Label: "On Balance (25%)"},
		},
	}
}

func loadConfig(path string) Config {
	cfg := defaultConfig()
	b, err := os.ReadFile(path)
	if err != nil {
		return cfg
	}
	var loaded Config
	if err := json.Unmarshal(b, &loaded); err != nil {
		return cfg
	}
	// Fill missing fields from defaults.
	if loaded.TaxYear != "" {
		cfg.TaxYear = loaded.TaxYear
	}
	if loaded.SalaryExemptionCap > 0 {
		cfg.SalaryExemptionCap = loaded.SalaryExemptionCap
	}
	if loaded.MinimumTax > 0 {
		cfg.MinimumTax = loaded.MinimumTax
	}
	if loaded.FestivalBonusRatio > 0 {
		cfg.FestivalBonusRatio = loaded.FestivalBonusRatio
	}
	if loaded.RebatePctOfTaxable > 0 {
		cfg.RebatePctOfTaxable = loaded.RebatePctOfTaxable
	}
	if loaded.RebatePctOfInvestment > 0 {
		cfg.RebatePctOfInvestment = loaded.RebatePctOfInvestment
	}
	if loaded.RebateMaxAmount > 0 {
		cfg.RebateMaxAmount = loaded.RebateMaxAmount
	}
	if loaded.SurchargeThreshold > 0 {
		cfg.SurchargeThreshold = loaded.SurchargeThreshold
	}
	if len(loaded.SurchargeAutoTiers) > 0 {
		cfg.SurchargeAutoTiers = loaded.SurchargeAutoTiers
	}
	if len(loaded.TaxSlabs) > 0 {
		cfg.TaxSlabs = loaded.TaxSlabs
	}
	return cfg
}

func (m model) Init() tea.Cmd {
	return textinput.Blink
}

func initialModel() model {
	cfg := loadConfig("config.json")

	s := spinner.New()
	s.Spinner = spinner.Dot
	s.Style = lipgloss.NewStyle().Foreground(lipgloss.Color("#0EA5E9"))

	// indexes mapping (0..30)
	inputLabels := []string{
		"1. Annual Salary / Total Package (BDT) [default: 0]",                        // 0
		"2. Festival bonus already included? (y/n) [default: n]",                     // 1
		"3. Enter custom salary breakdown? (y/n) [default: n]",                       // 2
		"   -> Basic pay (annual BDT) [default: 0]",                                  // 3
		"   -> House rent allowance (annual BDT) [default: 0]",                       // 4
		"   -> Medical allowance (annual BDT) [default: 0]",                          // 5
		"   -> Food allowance (annual BDT) [default: 0]",                             // 6
		"   -> Transport / conveyance (annual BDT) [default: 0]",                     // 7
		"   -> Mobile & other allowance (annual BDT) [default: 0]",                   // 8
		"   -> Festival bonus % of ONE MONTH BASIC (when custom) [default: 100%]",    // 9
		"4. Total annual expense (BDT) [default: 0]",                                 // 10
		"5. Location (dhaka/other) [default: other]",                                 // 11
		"6. Family size [default: 3]",                                                // 12
		"7. Do you have kids? (y/n) [default: n]",                                    // 13
		"8. Do you own your home? (y/n) [default: n]",                                // 14
		"9. Home-support staff? (y/n) [default: n]",                                  // 15
		"10. Mode (balanced/conservative/comfortable) [default: balanced]",           // 16
		"11. Previous year's gross income (BDT) [default: 0]",                        // 17
		"12. Net wealth (current) (BDT) [default: 0] (fallback)",                     // 18
		"13. Opening net wealth (previous year) (BDT) [default: 0]",                  // 19
		"14. Total assets (BDT) [default: 0] (optional; overrides net wealth input)", // 20
		"15. Bank loan outstanding (BDT) [default: 0]",                               // 21
		"16. Personal loan from others (BDT) [default: 0]",                           // 22
		"17. Credit card dues (BDT) [default: 0]",                                    // 23
		"18. Other liabilities (BDT) [default: 0]",                                   // 24
		"19. Apply net wealth surcharge? (y/n) [default: n]",                         // 25
		"20. Surcharge percent (number or 'auto') [default: auto]",                   // 26
		"21. Total eligible investment for rebate (BDT) [default: 0]",                // 27
		"22. Reverse calculator? (y/n) [default: n]",                                 // 28
		"23. Target net take-home pay (BDT) [default: 0]",                            // 29
		"24. Extra deductions / PF / other (BDT) [default: 0]",                       // 30
	}

	inps := make([]textinput.Model, len(inputLabels))
	for i := range inps {
		ti := textinput.New()
		ti.Placeholder = inputLabels[i]
		ti.CharLimit = 128
		ti.Width = 60
		if i == 0 {
			ti.Focus()
		}
		inps[i] = ti
	}

	prompt := textinput.New()
	prompt.Placeholder = "Enter file path..."
	prompt.CharLimit = 260
	prompt.Width = 60

	return model{
		state:   stateInput,
		step:    0,
		section: 0,
		inputs:  inps,
		spin:    s,
		width:   110,
		height:  40,
		config:  cfg,
		prompt:  prompt,
	}
}

func (m model) Update(msg tea.Msg) (tea.Model, tea.Cmd) {
	switch msg := msg.(type) {
	case tea.WindowSizeMsg:
		m.width = msg.Width
		m.height = msg.Height
		m.clampOffsets()
		return m, nil

	case spinner.TickMsg:
		if m.state == stateLoading {
			var cmd tea.Cmd
			m.spin, cmd = m.spin.Update(msg)
			return m, cmd
		}
		return m, nil

	case calcDoneMsg:
		m.state = stateResult
		m.report = msg.report
		m.resultOffset = 0
		m.notice = "Calculation complete."
		m.report.ReportText = formatReport(m.report)
		return m, nil

	case saveDoneMsg:
		m.notice = msg.message
		return m, nil

	case loadDoneMsg:
		m = msg.model
		m.notice = msg.message
		return m, nil

	case tea.KeyMsg:
		if m.promptActive {
			switch msg.String() {
			case "esc":
				m.promptActive = false
				m.state = stateResult
				return m, nil
			case "enter":
				path := strings.TrimSpace(m.prompt.Value())
				if path == "" {
					m.notice = "No path provided."
					m.promptActive = false
					m.state = stateResult
					return m, nil
				}
				m.promptActive = false
				m.state = stateResult
				return m, m.handlePromptAction(path)
			}
			var cmd tea.Cmd
			m.prompt, cmd = m.prompt.Update(msg)
			return m, cmd
		}

		switch msg.String() {
		case "ctrl+c", "q":
			return m, tea.Quit

		case "esc":
			if m.showInfo {
				m.showInfo = false
				return m, nil
			}
			if m.state == stateResult {
				m.state = stateInput
				return m, nil
			}
			return m, tea.Quit

		case "r":
			m = initialModel()
			return m, nil

		case "i":
			m.showInfo = !m.showInfo
			m.infoOffset = 0
			return m, nil

		case "ctrl+s":
			if m.state == stateResult {
				return m.openPrompt(promptSaveJSON, "tax_session.json", "Save session as JSON")
			}
		case "ctrl+p":
			if m.state == stateResult {
				return m.openPrompt(promptSavePDF, "tax_summary.pdf", "Save summary as PDF")
			}
		case "ctrl+e":
			if m.state == stateResult {
				return m.openPrompt(promptExportCSV, "tax_summary.csv", "Export summary as CSV")
			}
		case "ctrl+m":
			if m.state == stateResult {
				return m.openPrompt(promptExportMD, "tax_summary.md", "Export summary as Markdown")
			}
		case "ctrl+l":
			if m.state == stateResult || m.state == stateInput {
				return m.openPrompt(promptLoadJSON, "tax_session.json", "Load session JSON")
			}
		case "up", "k", "pgup", "home", "end", "down", "j", "pgdown":
			if m.state == stateResult && !m.showInfo {
				m.adjustResultScroll(msg.String())
				return m, nil
			}
			if m.showInfo {
				m.adjustInfoScroll(msg.String())
				return m, nil
			}

		case "tab", "enter":
			if m.state == stateInput {
				cmd := m.nextFieldOrCalculate()
				return m, cmd
			}

		case "backtab", "shift+tab":
			if m.state == stateInput {
				m.prevField()
				return m, nil
			}
		}

		if m.state == stateInput {
			var cmd tea.Cmd
			m.inputs[m.step], cmd = m.inputs[m.step].Update(msg)
			return m, cmd
		}
		return m, nil
	}

	return m, nil
}

func (m model) View() string {
	// Prompt overlay
	if m.promptActive {
		head := titleStyle.Render(" " + m.promptTitle() + " ")
		body := boxStyle.Width(maxInt(60, minInt(m.width-8, 80))).Render(
			"Location and filename:\n\n" + m.prompt.View() + "\n\n" + helpStyle.Render("Enter → confirm   Esc → cancel"),
		)
		return lipgloss.NewStyle().Margin(1, 2).Render(head + "\n\n" + body)
	}

	author := lipgloss.NewStyle().
		Bold(true).
		Foreground(lipgloss.Color("#FFFFFF")).
		Background(lipgloss.Color("#334155")).
		Padding(0, 1).
		Render(fmt.Sprintf(" %s • %s • %s ", appAuthor, appGitHub, time.Now().Format("2006-01-02")))

	banner := renderBanner(m.width)

	if m.state == stateLoading {
		return lipgloss.NewStyle().Margin(1, 2).Render(
			author + "\n\n" + banner + "\n\n" +
				boxStyle.Width(maxInt(60, minInt(m.width-8, 90))).Render(" "+m.spin.View()+" Computing tax, rebate, surcharge and report... "),
		)
	}

	if m.showInfo {
		details := renderDetailsPanel(m.width)
		visible := renderScrollableText(details, m.infoOffset, maxInt(8, m.height-10))
		footer := helpStyle.Render("Up/Down/PgUp/PgDn/Home/End → scroll   i → close details   q → quit")
		return wrapPage(author + "\n\n" + banner + "\n\n" + visible + "\n\n" + footer)
	}

	if m.state == stateResult {
		result := renderScrollableText(m.report.ReportText, m.resultOffset, maxInt(10, m.height-10))
		footer := helpStyle.Render("Up/Down/PgUp/PgDn/Home/End → scroll   ctrl+s JSON   ctrl+p PDF   ctrl+e CSV   ctrl+m MD   ctrl+l load   r restart   i details   q quit")
		topNotice := ""
		if m.notice != "" {
			topNotice = warnStyle.Render(" " + m.notice + " ")
		}
		return wrapPage(author + "\n\n" + banner + "\n\n" + topNotice + "\n\n" + result + "\n\n" + footer)
	}

	// Input screen
	var b strings.Builder
	b.WriteString(titleStyle.Render(" TAX COMPANION (BANGLADESH) ") + "\n\n")
	b.WriteString(sectionHeader(m.section) + "\n")

	for _, idx := range m.visibleSteps() {
		if idx > m.step {
			break
		}
		b.WriteString(keyStyle.Render(m.inputs[idx].Placeholder) + "\n")
		b.WriteString(m.inputs[idx].View() + "\n\n")
	}

	if m.notice != "" {
		b.WriteString(warnStyle.Render(m.notice) + "\n\n")
	}

	hint := helpStyle.Render("Enter/Tab → next   Shift+Tab → previous   ctrl+s save session   ctrl+p PDF   ctrl+e CSV   ctrl+m MD   ctrl+l load   r restart   i details   q quit")
	b.WriteString(hint)

	return wrapPage(author + "\n\n" + banner + "\n\n" + b.String())
}

func (m model) promptTitle() string {
	switch m.promptMode {
	case promptSavePDF:
		return "Save PDF Report"
	case promptSaveJSON:
		return "Save Session JSON"
	case promptExportCSV:
		return "Export CSV"
	case promptExportMD:
		return "Export Markdown"
	case promptLoadJSON:
		return "Load Session JSON"
	default:
		return "File"
	}
}

func wrapPage(s string) string {
	return lipgloss.NewStyle().Margin(1, 2).Render(s)
}

func sectionHeader(section int) string {
	switch section {
	case 0:
		return subTitleStyle.Render(" SECTION 1: INCOME ")
	case 1:
		return subTitleStyle.Render(" SECTION 2: EXPENSES ")
	case 2:
		return subTitleStyle.Render(" SECTION 3: WEALTH & REBATE ")
	case 3:
		return subTitleStyle.Render(" SECTION 4: REVERSE CALCULATOR ")
	default:
		return subTitleStyle.Render(" INPUTS ")
	}
}

func renderBanner(width int) string {
	// updated compact ASCII logo as requested
	lines := []string{
		"╭────────────────────────────────────────────────────────╮",
		"│  TAX COMPANION — IT-10B (Bangladesh)                   │",
		"│  Income, IT-10BB allocation, Wealth check & Audit risk │",
		"╰────────────────────────────────────────────────────────╯",
	}
	out := make([]string, 0, len(lines))
	maxw := maxPlainWidth(lines)
	for _, line := range lines {
		if width > maxw {
			pad := (width - maxw) / 2
			out = append(out, strings.Repeat(" ", pad)+line)
		} else {
			out = append(out, line)
		}
	}
	return strings.Join(out, "\n")
}

func maxPlainWidth(lines []string) int {
	mx := 0
	for _, line := range lines {
		if len(line) > mx {
			mx = len(line)
		}
	}
	return mx
}

func (m model) visibleSteps() []int {
	// updated indices covering new inputs
	steps := []int{0, 1, 2}
	if boolVal(m.inputs[2].Value(), false) {
		// custom breakdown: include breakdown fields 3..9 (festival percent included)
		steps = append(steps, 3, 4, 5, 6, 7, 8, 9)
	}
	// remaining inputs start at 10 .. 30
	steps = append(steps, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20, 21, 22, 23, 24, 25, 26, 27)
	// reverse calculator fields (28,29,30) are conditional on reverseEnabled (index 28)
	if boolVal(m.inputs[28].Value(), false) {
		steps = append(steps, 28, 29, 30)
	} else {
		steps = append(steps, 28)
	}
	return steps
}

func (m model) currentVisibleIndex() int {
	vis := m.visibleSteps()
	for i, v := range vis {
		if v == m.step {
			return i
		}
	}
	return 0
}

func (m *model) nextFieldOrCalculate() tea.Cmd {
	vis := m.visibleSteps()
	pos := m.currentVisibleIndex()
	if pos < len(vis)-1 {
		m.step = vis[pos+1]
		m.syncSectionAndFocus()
		return nil
	}
	m.state = stateLoading
	snapshot := *m
	for i := range snapshot.inputs {
		snapshot.inputs[i].Blur()
	}
	return tea.Batch(m.spin.Tick, computeReportCmd(snapshot))
}

func (m *model) prevField() {
	vis := m.visibleSteps()
	pos := m.currentVisibleIndex()
	if pos > 0 {
		m.step = vis[pos-1]
		m.syncSectionAndFocus()
	}
}

func (m *model) syncSectionAndFocus() {
	for i := range m.inputs {
		m.inputs[i].Blur()
	}
	if m.step >= 0 && m.step < len(m.inputs) {
		m.inputs[m.step].Focus()
	}
	m.section = stepToSection(m.step)
}

func stepToSection(step int) int {
	switch {
	case step <= 2:
		return 0
	case step <= 9:
		return 0
	case step <= 16:
		return 1
	case step <= 27:
		return 2
	default:
		return 3
	}
}

func (m *model) adjustResultScroll(key string) {
	switch key {
	case "up", "k":
		if m.resultOffset > 0 {
			m.resultOffset--
		}
	case "down", "j":
		m.resultOffset++
	case "pgup":
		m.resultOffset -= maxInt(1, m.height/2)
		if m.resultOffset < 0 {
			m.resultOffset = 0
		}
	case "pgdown":
		m.resultOffset += maxInt(1, m.height/2)
	case "home":
		m.resultOffset = 0
	case "end":
		m.resultOffset = 1 << 30
	}
	m.clampOffsets()
}

func (m *model) adjustInfoScroll(key string) {
	switch key {
	case "up", "k":
		if m.infoOffset > 0 {
			m.infoOffset--
		}
	case "down", "j":
		m.infoOffset++
	case "pgup":
		m.infoOffset -= maxInt(1, m.height/2)
		if m.infoOffset < 0 {
			m.infoOffset = 0
		}
	case "pgdown":
		m.infoOffset += maxInt(1, m.height/2)
	case "home":
		m.infoOffset = 0
	case "end":
		m.infoOffset = 1 << 30
	}
	m.clampOffsets()
}

func (m *model) clampOffsets() {
	lines := strings.Count(m.report.ReportText, "\n") + 1
	if lines < 1 {
		lines = 1
	}
	maxVisible := maxInt(6, m.height-10)
	if m.resultOffset > lines-maxVisible {
		m.resultOffset = maxInt(0, lines-maxVisible)
	}
	// details text length is stable enough to clamp conservatively
	if m.infoOffset < 0 {
		m.infoOffset = 0
	}
}

func (m model) handlePromptAction(path string) tea.Cmd {
	switch m.promptMode {
	case promptSavePDF:
		path = ensureExt(path, ".pdf")
		if err := exportPDF(path, m.report, m.config); err != nil {
			return func() tea.Msg { return saveDoneMsg{message: "PDF export failed: " + err.Error()} }
		}
		return func() tea.Msg { return saveDoneMsg{message: "PDF saved to " + path} }

	case promptSaveJSON:
		path = ensureExt(path, ".json")
		if err := saveSession(path, m); err != nil {
			return func() tea.Msg { return saveDoneMsg{message: "Save failed: " + err.Error()} }
		}
		return func() tea.Msg { return saveDoneMsg{message: "Session saved to " + path} }

	case promptExportCSV:
		path = ensureExt(path, ".csv")
		if err := exportCSV(path, m.report); err != nil {
			return func() tea.Msg { return saveDoneMsg{message: "CSV export failed: " + err.Error()} }
		}
		return func() tea.Msg { return saveDoneMsg{message: "CSV exported to " + path} }

	case promptExportMD:
		path = ensureExt(path, ".md")
		if err := exportMarkdown(path, m.report); err != nil {
			return func() tea.Msg { return saveDoneMsg{message: "Markdown export failed: " + err.Error()} }
		}
		return func() tea.Msg { return saveDoneMsg{message: "Markdown exported to " + path} }

	case promptLoadJSON:
		path = ensureExt(path, ".json")
		newM, err := loadSession(path, m.config)
		if err != nil {
			return func() tea.Msg { return saveDoneMsg{message: "Load failed: " + err.Error()} }
		}
		return func() tea.Msg { return loadDoneMsg{model: newM, message: "Session loaded from " + path} }
	}
	return nil
}

func computeReportCmd(snapshot model) tea.Cmd {
	return func() tea.Msg {
		time.Sleep(350 * time.Millisecond)
		return calcDoneMsg{report: buildReport(snapshot)}
	}
}

func buildReport(m model) Report {
	r := Report{GeneratedAt: time.Now(), TaxYear: m.config.TaxYear}

	// parse inputs with new indices mapping
	totalInput := parseNumeric(getVal(m.inputs[0].Value(), "0"))
	bonusIncluded := boolVal(m.inputs[1].Value(), false)
	customSalary := boolVal(m.inputs[2].Value(), false)

	basic := parseNumeric(getVal(m.inputs[3].Value(), "0"))
	hra := parseNumeric(getVal(m.inputs[4].Value(), "0"))
	med := parseNumeric(getVal(m.inputs[5].Value(), "0"))
	food := parseNumeric(getVal(m.inputs[6].Value(), "0"))
	trans := parseNumeric(getVal(m.inputs[7].Value(), "0"))
	mobile := parseNumeric(getVal(m.inputs[8].Value(), "0"))
	festivalPct := parseNumeric(getVal(m.inputs[9].Value(), "0")) // new: percent of ONE MONTH BASIC when custom

	totalExpense := parseNumeric(getVal(m.inputs[10].Value(), "0"))
	location := strings.ToLower(getVal(m.inputs[11].Value(), "other"))
	familySize := int(math.Round(parseNumeric(getVal(m.inputs[12].Value(), "3"))))
	if familySize <= 0 {
		familySize = 3
	}
	hasKids := boolVal(m.inputs[13].Value(), false)
	ownHome := boolVal(m.inputs[14].Value(), false)
	hasStaff := boolVal(m.inputs[15].Value(), false)
	mode := strings.ToLower(getVal(m.inputs[16].Value(), "balanced"))
	prevGross := parseNumeric(getVal(m.inputs[17].Value(), "0"))
	netWealthInput := parseNumeric(getVal(m.inputs[18].Value(), "0")) // fallback
	openingWealth := parseNumeric(getVal(m.inputs[19].Value(), "0"))
	totalAssets := parseNumeric(getVal(m.inputs[20].Value(), "0")) // new
	bankLoan := parseNumeric(getVal(m.inputs[21].Value(), "0"))
	personalLoan := parseNumeric(getVal(m.inputs[22].Value(), "0"))
	creditCard := parseNumeric(getVal(m.inputs[23].Value(), "0"))
	otherLiabilities := parseNumeric(getVal(m.inputs[24].Value(), "0"))
	applySurcharge := boolVal(m.inputs[25].Value(), false)
	surchargeMode := strings.ToLower(getVal(m.inputs[26].Value(), "auto"))
	investment := parseNumeric(getVal(m.inputs[27].Value(), "0"))
	reverseEnabled := boolVal(m.inputs[28].Value(), false)
	targetNet := parseNumeric(getVal(m.inputs[29].Value(), "0"))
	deductions := parseNumeric(getVal(m.inputs[30].Value(), "0"))

	r.SalaryInput = totalInput
	r.BonusIncluded = bonusIncluded
	r.CustomSalary = customSalary
	r.ReverseEnabled = reverseEnabled
	r.TargetNetTakeHome = targetNet
	r.Deductions = deductions
	r.RebateEligibleInvest = investment
	r.TotalExpense = totalExpense
	r.OpeningNetWealth = openingWealth
	r.ApplySurcharge = applySurcharge

	// festival / salary computations
	var baseGross float64
	var bonus float64
	var totalComp float64

	if customSalary {
		baseGross = roundTaka(basic + hra + med + food + trans + mobile)
		totalComp = roundTaka(totalInput)
		if totalComp <= 0 {
			totalComp = baseGross
		}
		// For custom breakdown festivalPct is percent of ONE MONTH BASIC.
		// Default: 100% (one full month) if user leaves it blank or zero.
		if festivalPct <= 0 {
			festivalPct = 100.0
		}
		if bonusIncluded {
			// If user indicated the totalComp already includes bonus, respect it.
			if totalComp < baseGross {
				bonus = 0
				totalComp = baseGross
			} else {
				bonus = totalComp - baseGross
			}
		} else {
			// festival bonus = (basic / 12) * (festivalPct / 100)
			monthlyBasic := basic / 12.0
			bonus = roundTaka(monthlyBasic * (festivalPct / 100.0))
			totalComp = baseGross + bonus
		}
	} else {
		// non-custom behaviour preserved
		if bonusIncluded {
			totalComp = roundTaka(totalInput)
			if totalComp <= 0 {
				totalComp = 0
			}
			// fallback: derive baseGross assuming config festival ratio
			baseGross = roundTaka(totalComp / (1 + m.config.FestivalBonusRatio))
			bonus = totalComp - baseGross
		} else {
			baseGross = roundTaka(totalInput)
			totalComp = baseGross
			if baseGross > 0 {
				// fallback festival bonus = baseGross * config ratio (legacy)
				bonus = roundTaka(baseGross * m.config.FestivalBonusRatio)
				totalComp = baseGross + bonus
			}
		}
	}

	r.BaseGrossSalary = baseGross
	r.HouseRentAllowance = hra
	r.MedicalAllowance = med
	r.ConveyanceAllowance = trans
	r.BonusAmount = bonus
	r.TotalComp = totalComp
	r.SalaryExempt = math.Min(totalComp/3.0, m.config.SalaryExemptionCap)
	r.TaxableSalary = math.Max(0, totalComp-r.SalaryExempt)

	r.CurrentTaxBeforeRebate, r.TaxLines = calculateTax(r.TaxableSalary, m.config.TaxSlabs, m.config.MinimumTax)
	r.RebateAmount = calculateRebate(r.TaxableSalary, investment, m.config, r.CurrentTaxBeforeRebate)
	r.CurrentTaxAfterRebate = math.Max(0, r.CurrentTaxBeforeRebate-r.RebateAmount)

	// Reverse calculator uses the current tax settings and deductions to approximate gross salary.
	if reverseEnabled && targetNet > 0 {
		r.EstimatedGrossFromNet = estimateGrossFromNet(targetNet, deductions, m, baseGross)
	}

	r.PreviousGrossInput = prevGross
	if prevGross > 0 {
		prevExempt := math.Min(prevGross/3.0, m.config.SalaryExemptionCap)
		prevTaxable := math.Max(0, prevGross-prevExempt)
		r.PreviousTax, r.PrevTaxLines = calculateTax(prevTaxable, m.config.TaxSlabs, m.config.MinimumTax)
	}

	r.CombinedBeforeSurch = r.CurrentTaxAfterRebate + r.PreviousTax

	// Liabilities & assets processing
	_ = creditCard
	_ = otherLiabilities
	_ = personalLoan
	totalLiabilities := bankLoan + personalLoan + creditCard + otherLiabilities
	r.TotalLiabilities = roundTaka(totalLiabilities)
	r.TotalAssets = roundTaka(totalAssets)

	// Decide which net wealth to use:
	netWealthUsed := netWealthInput
	if totalAssets > 0 {
		netWealthUsed = totalAssets - totalLiabilities
	}
	r.NetWealthCurrent = roundTaka(netWealthUsed)

	r.SurchargeRate = determineSurchargeRate(netWealthUsed, applySurcharge, surchargeMode, m.config)
	if r.SurchargeRate > 0 && netWealthUsed > m.config.SurchargeThreshold {
		r.SurchargeAmount = roundTaka(r.CombinedBeforeSurch * r.SurchargeRate)
	}
	r.FinalTax = r.CombinedBeforeSurch + r.SurchargeAmount

	// Expense allocation
	if totalExpense > 0 {
		r.ExpensePcts, r.ExpenseAmts = computeAllocation(totalExpense, location, familySize, hasKids, ownHome, hasStaff, mode)
	}

	r.WealthIncrease = r.NetWealthCurrent - r.OpeningNetWealth
	r.ExpectedSavings = totalComp - deductions - totalExpense - r.FinalTax
	r.WealthDifference = r.WealthIncrease - r.ExpectedSavings

	tol := math.Max(10000, math.Abs(r.ExpectedSavings)*0.01)
	switch {
	case math.Abs(r.WealthDifference) <= tol:
		r.WealthStatus = "OK — wealth increase matches estimated savings within tolerance."
	case r.WealthDifference > 0:
		r.WealthStatus = "ALERT — reported wealth is higher than estimated savings."
	default:
		r.WealthStatus = "Wealth increase is lower than estimated savings."
	}

	// Audit risk indicator (simple rule derived from wealth mismatch)
	if strings.HasPrefix(r.WealthStatus, "ALERT") {
		r.AuditRisk = "HIGH"
	} else if strings.HasPrefix(r.WealthStatus, "OK") {
		r.AuditRisk = "LOW"
	} else {
		r.AuditRisk = "MEDIUM"
	}

	r.ReportText = formatReport(r)
	return r
}

func estimateGrossFromNet(targetNet, deductions float64, m model, seedGross float64) float64 {
	// Binary search over gross salary. This is approximate but stable.
	low := 0.0
	high := math.Max(1000000, targetNet*3+1000000)
	if seedGross > high {
		high = seedGross * 2
	}

	best := 0.0
	bestDiff := math.MaxFloat64

	for i := 0; i < 80; i++ {
		mid := (low + high) / 2
		net := estimateNetFromGross(mid, deductions, m)
		diff := net - targetNet
		absDiff := math.Abs(diff)
		if absDiff < bestDiff {
			bestDiff = absDiff
			best = mid
		}
		if diff > 0 {
			high = mid
		} else {
			low = mid
		}
	}
	return roundTaka(best)
}

func estimateNetFromGross(baseGross, deductions float64, m model) float64 {
	// Use config festival ratio for non-custom estimate
	bonus := roundTaka(baseGross * m.config.FestivalBonusRatio)
	total := baseGross + bonus
	exempt := math.Min(total/3.0, m.config.SalaryExemptionCap)
	taxable := math.Max(0, total-exempt)
	tax, _ := calculateTax(taxable, m.config.TaxSlabs, m.config.MinimumTax)
	return total - tax - deductions
}

func determineSurchargeRate(netWealth float64, apply bool, mode string, cfg Config) float64 {
	if !apply || netWealth <= cfg.SurchargeThreshold {
		return 0
	}
	if mode == "" || mode == "auto" {
		for _, tier := range cfg.SurchargeAutoTiers {
			if netWealth <= tier.MaxNetWealth {
				return tier.Rate
			}
		}
		return 0
	}
	p := strings.TrimSuffix(mode, "%")
	v, err := strconv.ParseFloat(strings.TrimSpace(p), 64)
	if err != nil {
		return cfg.SurchargeAutoTiers[0].Rate
	}
	return v / 100.0
}

func calculateRebate(taxable, investment float64, cfg Config, taxBefore float64) float64 {
	if taxable <= 0 || investment <= 0 || taxBefore <= 0 {
		return 0
	}
	c1 := taxable * cfg.RebatePctOfTaxable
	c2 := investment * cfg.RebatePctOfInvestment
	rebate := math.Min(c1, math.Min(c2, cfg.RebateMaxAmount))
	rebate = roundTaka(rebate)
	if rebate > taxBefore {
		rebate = taxBefore
	}
	if rebate < 0 {
		rebate = 0
	}
	return rebate
}

func calculateTax(taxable float64, slabs []TaxSlab, minimumTax float64) (float64, []TaxLine) {
	remaining := taxable
	total := 0.0
	lines := make([]TaxLine, 0, len(slabs))

	for _, slab := range slabs {
		if remaining <= 0 {
			break
		}
		amt := remaining
		if slab.Limit > 0 && amt > slab.Limit {
			amt = slab.Limit
		}
		tax := amt * slab.Rate
		total += tax
		lines = append(lines, TaxLine{
			Label:  slab.Label,
			Amount: amt,
			Rate:   slab.Rate,
			Tax:    tax,
		})
		remaining -= amt
	}

	if total > 0 && total < minimumTax {
		lines = append(lines, TaxLine{
			Label:  "Minimum tax floor",
			Amount: 0,
			Rate:   0,
			Tax:    minimumTax - total,
		})
		total = minimumTax
	}
	return roundTaka(total), lines
}

func formatReport(r Report) string {
	var sb strings.Builder

	sb.WriteString(subTitleStyle.Render(" SUMMARY ") + "\n")
	sb.WriteString(fmt.Sprintf("%-34s %s\n", "Assessment year", r.TaxYear))
	sb.WriteString(fmt.Sprintf("%-34s %s\n", "Generated", r.GeneratedAt.Format("2006-01-02 15:04:05")))
	sb.WriteString("\n")

	sb.WriteString(subTitleStyle.Render(" SALARY / INCOME ") + "\n")
	sb.WriteString(kvLine("Annual salary input", r.SalaryInput))
	sb.WriteString(kvLine("House rent allowance", r.HouseRentAllowance))
	sb.WriteString(kvLine("Medical allowance", r.MedicalAllowance))
	sb.WriteString(kvLine("Conveyance allowance", r.ConveyanceAllowance))
	if r.CustomSalary {
		sb.WriteString(kvLine("Basic gross salary", r.BaseGrossSalary))
		sb.WriteString(kvLine("Festival bonus (one-month %)", r.BonusAmount))
	} else {
		sb.WriteString(kvLine("Base gross salary", r.BaseGrossSalary))
		sb.WriteString(kvLine("Festival bonus", r.BonusAmount))
	}
	sb.WriteString(kvLine("Total compensation", r.TotalComp))
	sb.WriteString(kvLine("Salary exemption", r.SalaryExempt))
	sb.WriteString(kvLine("Taxable income", r.TaxableSalary))
	sb.WriteString("\n")

	sb.WriteString(subTitleStyle.Render(" CURRENT YEAR TAX ") + "\n")
	sb.WriteString(taxTable(r.TaxLines))
	sb.WriteString(kvLine("Tax before rebate", r.CurrentTaxBeforeRebate))
	sb.WriteString(kvLine("Eligible investment", r.RebateEligibleInvest))
	sb.WriteString(kvLine("Tax rebate (Section 44)", r.RebateAmount))
	sb.WriteString(kvLine("Tax after rebate", r.CurrentTaxAfterRebate))
	sb.WriteString("\n")

	if r.PreviousGrossInput > 0 {
		sb.WriteString(subTitleStyle.Render(" PREVIOUS YEAR TAX ") + "\n")
		sb.WriteString(kvLine("Previous gross income", r.PreviousGrossInput))
		sb.WriteString(taxTable(r.PrevTaxLines))
		sb.WriteString(kvLine("Previous year tax", r.PreviousTax))
		sb.WriteString("\n")
	}

	if r.ApplySurcharge {
		sb.WriteString(subTitleStyle.Render(" NET WEALTH SURCHARGE ") + "\n")
		sb.WriteString(fmt.Sprintf("%-34s %s\n", "Applied", yesNo(r.SurchargeRate > 0)))
		sb.WriteString(fmt.Sprintf("%-34s %s\n", "Surcharge rate", fmt.Sprintf("%.2f%%", r.SurchargeRate*100)))
		sb.WriteString(kvLine("Surcharge amount", r.SurchargeAmount))
		sb.WriteString(kvLine("Combined tax before surcharge", r.CombinedBeforeSurch))
		sb.WriteString(kvLine("Final tax", r.FinalTax))
		sb.WriteString("\n")
	} else {
		sb.WriteString(subTitleStyle.Render(" FINAL TAX ") + "\n")
		sb.WriteString(kvLine("Final tax", r.FinalTax))
		sb.WriteString("\n")
	}

	if r.TotalExpense > 0 {
		sb.WriteString(subTitleStyle.Render(" IT-10BB EXPENSE ALLOCATION ") + "\n")
		keys := orderedExpenseKeys()
		for _, k := range keys {
			if r.ExpensePcts[k] == 0 && r.ExpenseAmts[k] == 0 {
				continue
			}
			sb.WriteString(fmt.Sprintf("%-34s %3d%%   Tk %s\n", k, r.ExpensePcts[k], formatMoney(r.ExpenseAmts[k])))
		}
		sb.WriteString(fmt.Sprintf("%-34s %3s   Tk %s\n", "TOTAL", "100%", formatMoney(int(roundTaka(r.TotalExpense)))))
		sb.WriteString("\n")
	}

	// Wealth summary & audit risk
	sb.WriteString(subTitleStyle.Render(" WEALTH SUMMARY ") + "\n")
	sb.WriteString(kvLine("Total assets (provided)", r.TotalAssets))
	sb.WriteString(kvLine("Total liabilities", r.TotalLiabilities))
	sb.WriteString(kvLine("Net wealth (used)", r.NetWealthCurrent))
	sb.WriteString(fmt.Sprintf("%-34s %s\n", "Audit risk", r.AuditRisk))
	sb.WriteString("\n")

	sb.WriteString(subTitleStyle.Render(" WEALTH CHECK ") + "\n")
	sb.WriteString(kvLine("Opening net wealth", r.OpeningNetWealth))
	sb.WriteString(kvLine("Wealth increase", r.WealthIncrease))
	sb.WriteString(kvLine("Estimated after-tax savings", r.ExpectedSavings))
	if r.WealthStatus != "" {
		if strings.HasPrefix(r.WealthStatus, "OK") {
			sb.WriteString(okStyle.Render(r.WealthStatus) + "\n")
		} else {
			sb.WriteString(warnStyle.Render(r.WealthStatus) + "\n")
		}
	}
	sb.WriteString("\n")

	if r.ReverseEnabled && r.TargetNetTakeHome > 0 {
		sb.WriteString(subTitleStyle.Render(" REVERSE CALCULATOR ") + "\n")
		sb.WriteString(kvLine("Target net take-home", r.TargetNetTakeHome))
		sb.WriteString(kvLine("Extra deductions", r.Deductions))
		sb.WriteString(kvLine("Estimated gross salary", r.EstimatedGrossFromNet))
		sb.WriteString("\n")
	}

	sb.WriteString(subTitleStyle.Render(" VISUAL SUMMARY ") + "\n")
	sb.WriteString(renderTaxPieChart(r.CurrentTaxAfterRebate, r.SurchargeAmount, 56) + "\n\n")
	if r.TotalExpense > 0 {
		sb.WriteString(renderExpensePieChart(r.ExpensePcts, r.ExpenseAmts, 56) + "\n\n")
	}
	sb.WriteString(helpStyle.Render("Tip: use ctrl+p for PDF, ctrl+s for session JSON, ctrl+e for CSV, ctrl+m for Markdown.") + "\n")

	return sb.String()
}

func taxTable(lines []TaxLine) string {
	if len(lines) == 0 {
		return "No taxable amount.\n"
	}
	var sb strings.Builder
	sb.WriteString(fmt.Sprintf("%-34s | %-12s | %-8s | %-12s\n", "SLAB", "TAXED", "RATE", "TAX"))
	sb.WriteString(strings.Repeat("-", 78) + "\n")
	for _, ln := range lines {
		sb.WriteString(fmt.Sprintf("%-34s | %12s | %7s | %12s\n",
			ln.Label,
			formatMoney(int(math.Round(ln.Amount))),
			fmt.Sprintf("%.0f%%", ln.Rate*100),
			formatMoney(int(math.Round(ln.Tax))),
		))
	}
	sb.WriteString(strings.Repeat("-", 78) + "\n")
	return sb.String()
}

func kvLine(label string, v float64) string {
	return fmt.Sprintf("%-34s %s\n", label, moneyStyle.Render("Tk "+formatMoney(int(roundTaka(v)))))
}

func yesNo(b bool) string {
	if b {
		return "Yes"
	}
	return "No"
}

func orderedExpenseKeys() []string {
	return []string{
		"Food, Clothing and Essentials",
		"Accommodation Expense",
		"Electricity",
		"Gas, Water, Sewer and Garbage",
		"Phone, Internet, TV & Subs",
		"Home-Support & Other Expenses",
		"Education Expenses",
		"Festival, Party, Events",
	}
}

func computeAllocation(total float64, loc string, familySize int, hasKids, ownHome, staff bool, mode string) (map[string]int, map[string]int) {
	weights := map[string]float64{
		"Food, Clothing and Essentials": 30.0,
		"Accommodation Expense":         28.0,
		"Electricity":                   2.5,
		"Gas, Water, Sewer and Garbage": 3.0,
		"Phone, Internet, TV & Subs":    3.5,
		"Home-Support & Other Expenses": 7.0,
		"Education Expenses":            10.0,
		"Festival, Party, Events":       6.0,
	}

	if !hasKids {
		weights["Education Expenses"] = 0
	}

	if strings.Contains(loc, "dhaka") || loc == "city" || loc == "metro" {
		weights["Accommodation Expense"] *= 1.20
		weights["Food, Clothing and Essentials"] *= 1.05
		weights["Home-Support & Other Expenses"] *= 1.10
	} else {
		weights["Accommodation Expense"] *= 0.90
	}

	extra := math.Max(0, float64(familySize-2))
	if extra > 0 {
		weights["Food, Clothing and Essentials"] *= 1 + 0.05*extra
		if hasKids {
			weights["Education Expenses"] *= 1 + 0.04*extra
		}
	}

	if ownHome {
		weights["Accommodation Expense"] *= 0.60
		weights["Home-Support & Other Expenses"] *= 1.05
	}
	if !staff {
		weights["Home-Support & Other Expenses"] *= 0.40
	}

	switch strings.ToLower(mode) {
	case "conservative":
		weights["Festival, Party, Events"] *= 0.60
		weights["Home-Support & Other Expenses"] *= 0.70
		weights["Food, Clothing and Essentials"] *= 1.08
	case "comfortable":
		weights["Festival, Party, Events"] *= 1.30
		weights["Home-Support & Other Expenses"] *= 1.20
		weights["Food, Clothing and Essentials"] *= 1.05
	}

	totalWeight := 0.0
	for _, v := range weights {
		totalWeight += v
	}
	if totalWeight <= 0 {
		return map[string]int{}, map[string]int{}
	}

	pcts := map[string]int{}
	remainders := make([]struct {
		key string
		rem float64
	}, 0, len(weights))
	sumPct := 0
	for k, v := range weights {
		raw := (v / totalWeight) * 100
		base := int(math.Floor(raw))
		pcts[k] = base
		sumPct += base
		remainders = append(remainders, struct {
			key string
			rem float64
		}{k, raw - float64(base)})
	}

	sort.Slice(remainders, func(i, j int) bool { return remainders[i].rem > remainders[j].rem })
	for i := 0; i < 100-sumPct; i++ {
		pcts[remainders[i%len(remainders)].key]++
	}

	amts := map[string]int{}
	amtRem := make([]struct {
		key string
		rem float64
	}, 0, len(pcts))
	totalInt := int(roundTaka(total))
	allocated := 0
	for k, p := range pcts {
		rawAmt := float64(totalInt) * float64(p) / 100
		base := int(math.Floor(rawAmt))
		amts[k] = base
		allocated += base
		amtRem = append(amtRem, struct {
			key string
			rem float64
		}{k, rawAmt - float64(base)})
	}

	drift := totalInt - allocated
	if drift > 0 {
		sort.Slice(amtRem, func(i, j int) bool { return amtRem[i].rem > amtRem[j].rem })
		for i := 0; i < drift; i++ {
			amts[amtRem[i%len(amtRem)].key]++
		}
	} else if drift < 0 {
		sort.Slice(amtRem, func(i, j int) bool { return amtRem[i].rem < amtRem[j].rem })
		for i := 0; i < -drift; i++ {
			amts[amtRem[i%len(amtRem)].key]--
		}
	}

	return pcts, amts
}

func renderPieChart(title string, items []TaxLine, width int) string {
	filtered := make([]TaxLine, 0, len(items))
	for _, it := range items {
		if it.Tax > 0 {
			filtered = append(filtered, it)
		}
	}
	if len(filtered) == 0 {
		return title + "\nNo data to display."
	}
	sort.Slice(filtered, func(i, j int) bool { return filtered[i].Tax > filtered[j].Tax })
	total := 0.0
	for _, it := range filtered {
		total += it.Tax
	}
	if total <= 0 {
		return title + "\nNo data to display."
	}

	segmentCount := 24
	bar := strings.Builder{}
	bar.WriteString(title + "\n")
	center := ""
	for i := 0; i < segmentCount; i++ {
		center += "●"
	}
	if width > len(center) {
		bar.WriteString(strings.Repeat(" ", (width-len(center))/2))
	}
	bar.WriteString(center + "\n\n")
	for _, it := range filtered {
		pct := (it.Tax / total) * 100
		bar.WriteString(fmt.Sprintf("• %-28s | %3.0f%% | Tk %s\n", it.Label, pct, formatMoney(int(roundTaka(it.Tax)))))
	}
	return bar.String()
}

func renderTaxPieChart(tax, surcharge float64, width int) string {
	items := []TaxLine{{Label: "Current tax", Tax: tax}}
	if surcharge > 0 {
		items = append(items, TaxLine{Label: "Surcharge", Tax: surcharge})
	}
	return renderPieChart("Tax Composition", items, width)
}

func renderExpensePieChart(pcts map[string]int, amts map[string]int, width int) string {
	items := make([]TaxLine, 0, len(pcts))
	for _, k := range orderedExpenseKeys() {
		if pcts[k] > 0 {
			items = append(items, TaxLine{Label: k, Tax: float64(amts[k])})
		}
	}
	return renderPieChart("Expense Allocation", items, width)
}

func renderScrollableText(content string, offset, height int) string {
	lines := strings.Split(content, "\n")
	if height < 6 {
		height = 6
	}
	if offset < 0 {
		offset = 0
	}
	if offset > len(lines)-height {
		offset = maxInt(0, len(lines)-height)
	}
	end := offset + height
	if end > len(lines) {
		end = len(lines)
	}
	return strings.Join(lines[offset:end], "\n")
}

func renderDetailsPanel(width int) string {
	w := width - 8
	if w < 48 {
		w = 48
	}
	var sb strings.Builder
	sb.WriteString(subTitleStyle.Render(" INPUT GUIDE ") + "\n\n")
	write := func(k, d string) {
		sb.WriteString(keyStyle.Render(k) + "\n")
		sb.WriteString(wrapText(d, w) + "\n\n")
	}
	write("Annual Salary / Total Package", "Enter gross annual salary or the total package amount. Expressions like 12809*23 and shorthands like 1 lakh, 2cr, 4.5k are accepted.")
	write("Festival bonus already included?", "Use y if the salary figure already includes festival bonus. Use n if you want the tool to estimate the bonus separately.")
	write("Custom salary breakdown?", "Enable if you want to provide exact salary components such as basic, house rent, medical, food, transport and mobile allowances.")
	write("Festival bonus % (when custom)", "Festival bonus is now calculated as a percent of ONE MONTH BASIC: (basic / 12) * (pct/100). Default is 100% (one full month's basic).")
	write("Total annual expense", "Used for IT-10BB style allocation. The tool distributes every taka into expense buckets without drift.")
	write("Location / family / home / staff / mode", "Used to adjust household allocation weights. Dhaka increases housing weight; conservative mode reduces festival and support expenses.")
	write("Investment for rebate", "Eligible investment amount for the Bangladesh general tax rebate. The tool applies the lower of 3% of taxable income, 15% of investment, and Tk 10 lakh by default.")
	write("Net wealth & liabilities", "Provide net wealth directly (fallback) or enter Total Assets plus liabilities (bank loans, personal loans, credit card dues, other liabilities). If Total Assets is provided it overrides manual net wealth input and computes Net Wealth = Total Assets - Total Liabilities.")
	write("Reverse calculator", "If enabled, you can enter a target net take-home amount. The tool estimates the gross salary needed to reach that net after tax and deductions.")
	write("Save / load", "Ctrl+S saves the current session as JSON with SHA256 integrity. Ctrl+L loads it back after verifying the hash.")
	write("PDF / CSV / Markdown export", "Ctrl+P exports a professional PDF including your logo.png (if present) and a watermark. Ctrl+E exports CSV. Ctrl+M exports Markdown.")
	return boxStyle.Width(w).Render(sb.String())
}

func wrapText(s string, width int) string {
	if width < 20 {
		width = 20
	}
	words := strings.Fields(s)
	if len(words) == 0 {
		return ""
	}
	var lines []string
	cur := words[0]
	for _, w := range words[1:] {
		if len(cur)+1+len(w) > width {
			lines = append(lines, cur)
			cur = w
		} else {
			cur += " " + w
		}
	}
	lines = append(lines, cur)
	return strings.Join(lines, "\n")
}

func saveSession(path string, m model) error {
	session := SessionData{
		Version:   1,
		Timestamp: time.Now().Format(time.RFC3339),
		Inputs:    currentValues(m.inputs),
		Step:      m.step,
		Section:   m.section,
	}
	raw, err := json.Marshal(session)
	if err != nil {
		return err
	}
	hash := sha256.Sum256(raw)
	wrapped := SavedFile{
		Hash:    fmt.Sprintf("%x", hash[:]),
		Session: session,
	}
	out, err := json.MarshalIndent(wrapped, "", "  ")
	if err != nil {
		return err
	}
	return os.WriteFile(path, out, 0644)
}

func loadSession(path string, cfg Config) (model, error) {
	b, err := os.ReadFile(path)
	if err != nil {
		return model{}, err
	}
	var wrapped SavedFile
	if err := json.Unmarshal(b, &wrapped); err != nil {
		return model{}, err
	}
	raw, err := json.Marshal(wrapped.Session)
	if err != nil {
		return model{}, err
	}
	check := sha256.Sum256(raw)
	if fmt.Sprintf("%x", check[:]) != wrapped.Hash {
		return model{}, fmt.Errorf("SHA256 mismatch: file may be corrupted or tampered with")
	}
	m := initialModel()
	m.config = cfg
	for i := range m.inputs {
		if i < len(wrapped.Session.Inputs) {
			m.inputs[i].SetValue(wrapped.Session.Inputs[i])
		}
	}
	m.step = wrapped.Session.Step
	if m.step < 0 || m.step >= len(m.inputs) {
		m.step = 0
	}
	m.section = wrapped.Session.Section
	m.syncSectionAndFocus()
	m.state = stateInput
	m.notice = "Session loaded successfully."
	return m, nil
}

func exportCSV(path string, r Report) error {
	f, err := os.Create(path)
	if err != nil {
		return err
	}
	defer f.Close()

	w := csv.NewWriter(f)
	defer w.Flush()

	rows := [][]string{
		{"Field", "Value"},
		{"Tax year", r.TaxYear},
		{"Annual salary input", formatMoney(int(roundTaka(r.SalaryInput)))},
		{"Base gross salary", formatMoney(int(roundTaka(r.BaseGrossSalary)))},
		{"House rent allowance", formatMoney(int(roundTaka(r.HouseRentAllowance)))},
		{"Medical allowance", formatMoney(int(roundTaka(r.MedicalAllowance)))},
		{"Conveyance allowance", formatMoney(int(roundTaka(r.ConveyanceAllowance)))},
		{"Festival bonus", formatMoney(int(roundTaka(r.BonusAmount)))},
		{"Total compensation", formatMoney(int(roundTaka(r.TotalComp)))},
		{"Salary exemption", formatMoney(int(roundTaka(r.SalaryExempt)))},
		{"Taxable income", formatMoney(int(roundTaka(r.TaxableSalary)))},
		{"Tax before rebate", formatMoney(int(roundTaka(r.CurrentTaxBeforeRebate)))},
		{"Rebate eligible investment", formatMoney(int(roundTaka(r.RebateEligibleInvest)))},
		{"Tax rebate", formatMoney(int(roundTaka(r.RebateAmount)))},
		{"Tax after rebate", formatMoney(int(roundTaka(r.CurrentTaxAfterRebate)))},
		{"Previous year tax", formatMoney(int(roundTaka(r.PreviousTax)))},
		{"Surcharge amount", formatMoney(int(roundTaka(r.SurchargeAmount)))},
		{"Final tax", formatMoney(int(roundTaka(r.FinalTax)))},
		{"Total assets (provided)", formatMoney(int(roundTaka(r.TotalAssets)))},
		{"Total liabilities", formatMoney(int(roundTaka(r.TotalLiabilities)))},
		{"Current net wealth", formatMoney(int(roundTaka(r.NetWealthCurrent)))},
		{"Opening net wealth", formatMoney(int(roundTaka(r.OpeningNetWealth)))},
		{"Wealth increase", formatMoney(int(roundTaka(r.WealthIncrease)))},
		{"Estimated after-tax savings", formatMoney(int(roundTaka(r.ExpectedSavings)))},
		{"Wealth status", r.WealthStatus},
		{"Audit risk", r.AuditRisk},
	}
	for _, row := range rows {
		if err := w.Write(row); err != nil {
			return err
		}
	}

	if r.TotalExpense > 0 {
		_ = w.Write([]string{})
		_ = w.Write([]string{"Expense category", "Percentage", "Amount"})
		for _, k := range orderedExpenseKeys() {
			if r.ExpensePcts[k] == 0 && r.ExpenseAmts[k] == 0 {
				continue
			}
			_ = w.Write([]string{k, fmt.Sprintf("%d", r.ExpensePcts[k]), formatMoney(r.ExpenseAmts[k])})
		}
	}
	return nil
}

func exportMarkdown(path string, r Report) error {
	var sb strings.Builder
	sb.WriteString("# Tax Companion Report\n\n")
	sb.WriteString(fmt.Sprintf("* Tax year: %s\n", r.TaxYear))
	sb.WriteString(fmt.Sprintf("* Generated: %s\n\n", r.GeneratedAt.Format(time.RFC3339)))

	sb.WriteString("## Summary\n\n")
	sb.WriteString("| Field | Value |\n|---|---:|\n")
	sb.WriteString(fmt.Sprintf("| Annual salary input | Tk %s |\n", formatMoney(int(roundTaka(r.SalaryInput)))))
	sb.WriteString(fmt.Sprintf("| Base gross salary | Tk %s |\n", formatMoney(int(roundTaka(r.BaseGrossSalary)))))
	sb.WriteString(fmt.Sprintf("| House rent allowance | Tk %s |\n", formatMoney(int(roundTaka(r.HouseRentAllowance)))))
	sb.WriteString(fmt.Sprintf("| Medical allowance | Tk %s |\n", formatMoney(int(roundTaka(r.MedicalAllowance)))))
	sb.WriteString(fmt.Sprintf("| Conveyance allowance | Tk %s |\n", formatMoney(int(roundTaka(r.ConveyanceAllowance)))))
	sb.WriteString(fmt.Sprintf("| Festival bonus | Tk %s |\n", formatMoney(int(roundTaka(r.BonusAmount)))))
	sb.WriteString(fmt.Sprintf("| Total compensation | Tk %s |\n", formatMoney(int(roundTaka(r.TotalComp)))))
	sb.WriteString(fmt.Sprintf("| Salary exemption | Tk %s |\n", formatMoney(int(roundTaka(r.SalaryExempt)))))
	sb.WriteString(fmt.Sprintf("| Taxable income | Tk %s |\n", formatMoney(int(roundTaka(r.TaxableSalary)))))
	sb.WriteString(fmt.Sprintf("| Tax before rebate | Tk %s |\n", formatMoney(int(roundTaka(r.CurrentTaxBeforeRebate)))))
	sb.WriteString(fmt.Sprintf("| Eligible investment | Tk %s |\n", formatMoney(int(roundTaka(r.RebateEligibleInvest)))))
	sb.WriteString(fmt.Sprintf("| Tax rebate | Tk %s |\n", formatMoney(int(roundTaka(r.RebateAmount)))))
	sb.WriteString(fmt.Sprintf("| Tax after rebate | Tk %s |\n", formatMoney(int(roundTaka(r.CurrentTaxAfterRebate)))))
	sb.WriteString(fmt.Sprintf("| Previous year tax | Tk %s |\n", formatMoney(int(roundTaka(r.PreviousTax)))))
	sb.WriteString(fmt.Sprintf("| Surcharge amount | Tk %s |\n", formatMoney(int(roundTaka(r.SurchargeAmount)))))
	sb.WriteString(fmt.Sprintf("| Final tax | Tk %s |\n", formatMoney(int(roundTaka(r.FinalTax)))))
	sb.WriteString(fmt.Sprintf("| Total assets | Tk %s |\n", formatMoney(int(roundTaka(r.TotalAssets)))))
	sb.WriteString(fmt.Sprintf("| Total liabilities | Tk %s |\n", formatMoney(int(roundTaka(r.TotalLiabilities)))))
	sb.WriteString(fmt.Sprintf("| Net wealth (used) | Tk %s |\n", formatMoney(int(roundTaka(r.NetWealthCurrent)))))
	sb.WriteString(fmt.Sprintf("| Opening net wealth | Tk %s |\n", formatMoney(int(roundTaka(r.OpeningNetWealth)))))
	sb.WriteString(fmt.Sprintf("| Wealth increase | Tk %s |\n", formatMoney(int(roundTaka(r.WealthIncrease)))))
	sb.WriteString(fmt.Sprintf("| Estimated after-tax savings | Tk %s |\n", formatMoney(int(roundTaka(r.ExpectedSavings)))))
	sb.WriteString(fmt.Sprintf("| Wealth status | %s |\n", r.WealthStatus))
	sb.WriteString(fmt.Sprintf("| Audit risk | %s |\n\n", r.AuditRisk))

	if r.TotalExpense > 0 {
		sb.WriteString("## Expense allocation\n\n")
		sb.WriteString("| Category | % | Amount |\n|---|---:|---:|\n")
		for _, k := range orderedExpenseKeys() {
			if r.ExpensePcts[k] == 0 && r.ExpenseAmts[k] == 0 {
				continue
			}
			sb.WriteString(fmt.Sprintf("| %s | %d | Tk %s |\n", k, r.ExpensePcts[k], formatMoney(r.ExpenseAmts[k])))
		}
		sb.WriteString("\n")
	}

	sb.WriteString("## Wealth check\n\n")
	sb.WriteString(fmt.Sprintf("* Total assets: Tk %s\n", formatMoney(int(roundTaka(r.TotalAssets)))))
	sb.WriteString(fmt.Sprintf("* Total liabilities: Tk %s\n", formatMoney(int(roundTaka(r.TotalLiabilities)))))
	sb.WriteString(fmt.Sprintf("* Net wealth (used): Tk %s\n", formatMoney(int(roundTaka(r.NetWealthCurrent)))))
	sb.WriteString(fmt.Sprintf("* Audit risk: %s\n\n", r.AuditRisk))

	if r.ReverseEnabled && r.TargetNetTakeHome > 0 {
		sb.WriteString("## Reverse calculator\n\n")
		sb.WriteString(fmt.Sprintf("* Target net take-home: Tk %s\n", formatMoney(int(roundTaka(r.TargetNetTakeHome)))))
		sb.WriteString(fmt.Sprintf("* Estimated gross salary: Tk %s\n", formatMoney(int(roundTaka(r.EstimatedGrossFromNet)))))
	}
	return os.WriteFile(path, []byte(sb.String()), 0644)
}

func exportPDF(path string, r Report, cfg Config) error {
	pdf := gofpdf.New("P", "mm", "A4", "")
	pdf.SetMargins(12, 12, 12)
	pdf.SetAutoPageBreak(true, 12)
	pdf.AddPage()

	// Header: try to use logo.png in current directory (expected 124x124 px).
	// Convert 124px @ 96 DPI -> mm: (124/96)*25.4 ≈ 32.75 mm
	const margin = 12.0
	logoSizeMM := (124.0 / 96.0) * 25.4
	logoPath := "logo.png"
	headerTop := margin
	headerTextX := margin + logoSizeMM + 8
	hasLogo := false
	if _, err := os.Stat(logoPath); err == nil {
		hasLogo = true
		// place logo at left margin
		pdf.ImageOptions(logoPath, margin, margin, logoSizeMM, logoSizeMM, false, gofpdf.ImageOptions{ImageType: "", ReadDpi: true}, 0, "")
	} else {
		// fallback placeholder square
		rC, gC, bC := hexToRGB(theme.Primary)
		pdf.SetDrawColor(rC, gC, bC)
		pdf.SetFillColor(rC, gC, bC)
		pdf.Rect(margin, margin, 22, 22, "DF")
		pdf.SetTextColor(255, 255, 255)
		pdf.SetFont("Arial", "B", 14)
		pdf.SetXY(margin+4, margin+6)
		pdf.CellFormat(14, 6, "TC", "", 0, "C", false, 0, "")
		pdf.SetTextColor(0, 0, 0)
		headerTextX = margin + 28
	}

	pdf.SetFont("Arial", "B", 18)
	pdf.SetXY(headerTextX, headerTop+1)
	pdf.CellFormat(0, 8, "Tax Companion Report", "", 0, "L", false, 0, "")
	pdf.SetFont("Arial", "", 10)
	pdf.SetXY(headerTextX, headerTop+9)
	pdf.CellFormat(0, 5, fmt.Sprintf("Prepared for: %s   |   Author: %s   |   GitHub: %s", appName, appAuthor, appGitHub), "", 0, "L", false, 0, "")
	pdf.SetXY(headerTextX, headerTop+14)
	pdf.CellFormat(0, 5, fmt.Sprintf("Assessment year: %s   |   Generated: %s", r.TaxYear, r.GeneratedAt.Format("2006-01-02 15:04:05")), "", 0, "L", false, 0, "")

	headerBottom := headerTop + 14 + 5
	if hasLogo {
		if candidate := headerTop + logoSizeMM; candidate > headerBottom {
			headerBottom = candidate
		}
	} else {
		if candidate := headerTop + 22; candidate > headerBottom {
			headerBottom = candidate
		}
	}
	pdf.SetY(headerBottom + 3)
	drawRule(pdf)

	section := func(title string) {
		pdf.Ln(1)
		r2, g2, b2 := hexToRGB(theme.Secondary)
		pdf.SetFillColor(r2, g2, b2)
		rC, gC, bC := hexToRGB(theme.Primary)
		pdf.SetTextColor(rC, gC, bC)
		pdf.SetFont("Arial", "B", 12)
		pdf.CellFormat(0, 8, " "+title+" ", "1", 1, "L", true, 0, "")
		pdf.SetTextColor(0, 0, 0)
		pdf.SetFont("Arial", "", 10)
	}

	kv := func(label string, value string) {
		pdf.CellFormat(70, 6, label, "", 0, "L", false, 0, "")
		pdf.CellFormat(0, 6, value, "", 1, "L", false, 0, "")
	}

	section("Salary / Income")
	kv("Annual salary input", "Tk "+formatMoney(int(roundTaka(r.SalaryInput))))
	kv("Base gross salary", "Tk "+formatMoney(int(roundTaka(r.BaseGrossSalary))))
	kv("House rent allowance", "Tk "+formatMoney(int(roundTaka(r.HouseRentAllowance))))
	kv("Medical allowance", "Tk "+formatMoney(int(roundTaka(r.MedicalAllowance))))
	kv("Conveyance allowance", "Tk "+formatMoney(int(roundTaka(r.ConveyanceAllowance))))
	kv("Festival bonus", "Tk "+formatMoney(int(roundTaka(r.BonusAmount))))
	kv("Total compensation", "Tk "+formatMoney(int(roundTaka(r.TotalComp))))
	kv("Salary exemption", "Tk "+formatMoney(int(roundTaka(r.SalaryExempt))))
	kv("Taxable income", "Tk "+formatMoney(int(roundTaka(r.TaxableSalary))))

	section("Current Year Tax")
	for _, ln := range r.TaxLines {
		pdf.CellFormat(110, 6, ln.Label, "", 0, "L", false, 0, "")
		pdf.CellFormat(0, 6, fmt.Sprintf("Tk %s", formatMoney(int(roundTaka(ln.Tax)))), "", 1, "R", false, 0, "")
	}
	kv("Tax before rebate", "Tk "+formatMoney(int(roundTaka(r.CurrentTaxBeforeRebate))))
	kv("Eligible investment", "Tk "+formatMoney(int(roundTaka(r.RebateEligibleInvest))))
	kv("Tax rebate", "Tk "+formatMoney(int(roundTaka(r.RebateAmount))))
	kv("Tax after rebate", "Tk "+formatMoney(int(roundTaka(r.CurrentTaxAfterRebate))))

	if r.PreviousGrossInput > 0 {
		section("Previous Year Tax")
		kv("Previous gross income", "Tk "+formatMoney(int(roundTaka(r.PreviousGrossInput))))
		kv("Previous year tax", "Tk "+formatMoney(int(roundTaka(r.PreviousTax))))
	}

	if r.ApplySurcharge {
		section("Net Wealth Surcharge")
		kv("Surcharge rate", fmt.Sprintf("%.2f%%", r.SurchargeRate*100))
		kv("Surcharge amount", "Tk "+formatMoney(int(roundTaka(r.SurchargeAmount))))
		kv("Combined before surcharge", "Tk "+formatMoney(int(roundTaka(r.CombinedBeforeSurch))))
		kv("Final tax", "Tk "+formatMoney(int(roundTaka(r.FinalTax))))
	} else {
		section("Final Tax")
		kv("Final tax", "Tk "+formatMoney(int(roundTaka(r.FinalTax))))
	}

	if r.TotalExpense > 0 {
		section("IT-10BB Expense Allocation")
		for _, k := range orderedExpenseKeys() {
			if r.ExpensePcts[k] == 0 && r.ExpenseAmts[k] == 0 {
				continue
			}
			pdf.CellFormat(92, 6, k, "", 0, "L", false, 0, "")
			pdf.CellFormat(0, 6, fmt.Sprintf("%d%%   Tk %s", r.ExpensePcts[k], formatMoney(r.ExpenseAmts[k])), "", 1, "R", false, 0, "")
		}
	}

	section("Wealth Summary")
	kv("Total assets (provided)", "Tk "+formatMoney(int(roundTaka(r.TotalAssets))))
	kv("Total liabilities", "Tk "+formatMoney(int(roundTaka(r.TotalLiabilities))))
	kv("Net wealth (used)", "Tk "+formatMoney(int(roundTaka(r.NetWealthCurrent))))
	kv("Audit risk", r.AuditRisk)
	kv("Wealth increase", "Tk "+formatMoney(int(roundTaka(r.WealthIncrease))))
	kv("Estimated after-tax savings", "Tk "+formatMoney(int(roundTaka(r.ExpectedSavings))))
	kv("Status", r.WealthStatus)

	section("Summary")
	pdf.MultiCell(0, 5, "This report summarizes salary, tax, rebate, surcharge, expense allocation, wealth consistency and reverse salary estimation. Values depend on the inputs provided in the TUI and the tax slab configuration.", "", "L", false)

	// Add watermark on every page
	totalPages := pdf.PageNo()
	for p := 1; p <= totalPages; p++ {
		pdf.SetPage(p)
		addWatermark(pdf, "Companion by rhshourav")
	}

	return pdf.OutputFileAndClose(path)
}

// addWatermark draws a light rotated watermark text on the page center.

func addWatermark(pdf *gofpdf.Fpdf, text string) {
	// Use theme values
	pdf.SetFont("Arial", "B", theme.WatermarkSize)
	// light gray as default
	r, g, b := 200, 200, 200
	// let primary color influence watermark slightly if theme primary provided
	if theme.Primary != "" {
		r2, g2, b2 := hexToRGB(theme.Primary)
		// pick a lighter variant for watermark
		r = (r2 + 255) / 2
		g = (g2 + 255) / 2
		b = (b2 + 255) / 2
	}
	pdf.SetTextColor(r, g, b)
	// semi-transparent
	pdf.SetAlpha(theme.WatermarkOpacity, "Normal")
	// Get page dimensions
	pageW, pageH := pdf.GetPageSize()
	cx := pageW / 2.0
	cy := pageH / 2.0
	// Rotate around center and write text
	pdf.TransformBegin()
	pdf.TransformRotate(theme.WatermarkAngle, cx, cy)
	wd := pdf.GetStringWidth(text)
	x := cx - wd/2.0
	y := cy
	pdf.Text(x, y, text)
	pdf.TransformEnd()
	// reset alpha to full
	pdf.SetAlpha(1.0, "Normal")
	// reset color to black
	pdf.SetTextColor(0, 0, 0)
}

func drawRule(pdf *gofpdf.Fpdf) {
	x1, y := pdf.GetX(), pdf.GetY()
	_ = x1
	rC, gC, bC := hexToRGB(theme.Primary)
	pdf.SetDrawColor(rC, gC, bC)
	pdf.SetLineWidth(0.4)
	pdf.Line(12, y, 198, y)
	pdf.Ln(3)
}

func getVal(v, def string) string {
	s := strings.TrimSpace(v)
	if s == "" {
		return def
	}
	return s
}

func boolVal(v string, def bool) bool {
	s := strings.ToLower(strings.TrimSpace(v))
	if s == "" {
		return def
	}
	switch s {
	case "y", "yes", "true", "1", "on":
		return true
	case "n", "no", "false", "0", "off":
		return false
	default:
		return def
	}
}

func parseNumeric(s string) float64 {
	s = strings.TrimSpace(s)
	if s == "" {
		return 0
	}
	s = strings.ToLower(strings.ReplaceAll(s, ",", ""))

	replacements := []struct {
		re   *regexp.Regexp
		repl string
	}{
		{regexp.MustCompile(`(?i)(\d+(?:\.\d+)?)\s*(k|thousand)\b`), "$1*1000"},
		{regexp.MustCompile(`(?i)(\d+(?:\.\d+)?)\s*(m|mn|million)\b`), "$1*1000000"},
		{regexp.MustCompile(`(?i)(\d+(?:\.\d+)?)\s*(lakh|lac|lacs)\b`), "$1*100000"},
		{regexp.MustCompile(`(?i)(\d+(?:\.\d+)?)\s*(cr|crore|crores)\b`), "$1*10000000"},
		{regexp.MustCompile(`(?i)(\d+(?:\.\d+)?)\s*%`), "$1/100"},
	}
	for _, r := range replacements {
		s = r.re.ReplaceAllString(s, r.repl)
	}

	// evaluate simple expressions like 12809*23
	val, err := evalSimpleExpression(s)
	if err != nil {
		return 0
	}
	return val
}

func evalSimpleExpression(expr string) (float64, error) {
	// very small evaluator supporting + - * / and parentheses, integers and floats
	expr = strings.TrimSpace(expr)
	if expr == "" {
		return 0, nil
	}
	toks := tokenize(expr)
	ev, err := parseExpr(toks)
	if err != nil {
		return 0, err
	}
	if len(ev) != 1 {
		return 0, fmt.Errorf("bad expression")
	}
	return ev[0], nil
}

func tokenize(s string) []string {
	var out []string
	cur := ""
	for _, r := range s {
		switch {
		case r == '+' || r == '-' || r == '*' || r == '/' || r == '(' || r == ')':
			if cur != "" {
				out = append(out, cur)
				cur = ""
			}
			out = append(out, string(r))
		case r == ' ' || r == '\t' || r == '\n':
			if cur != "" {
				out = append(out, cur)
				cur = ""
			}
		default:
			cur += string(r)
		}
	}
	if cur != "" {
		out = append(out, cur)
	}
	return out
}

func parseExpr(tokens []string) ([]float64, error) {
	// Shunting-yard-lite using recursive descent for simplicity
	var stackNums []float64
	var stackOps []string

	applyOp := func() error {
		if len(stackNums) < 2 || len(stackOps) == 0 {
			return nil
		}
		b := stackNums[len(stackNums)-1]
		a := stackNums[len(stackNums)-2]
		op := stackOps[len(stackOps)-1]
		stackNums = stackNums[:len(stackNums)-2]
		stackOps = stackOps[:len(stackOps)-1]
		switch op {
		case "+":
			stackNums = append(stackNums, a+b)
		case "-":
			stackNums = append(stackNums, a-b)
		case "*":
			stackNums = append(stackNums, a*b)
		case "/":
			if b == 0 {
				return fmt.Errorf("divide by zero")
			}
			stackNums = append(stackNums, a/b)
		default:
			return fmt.Errorf("unknown op")
		}
		return nil
	}

	prec := func(op string) int {
		if op == "+" || op == "-" {
			return 1
		}
		if op == "*" || op == "/" {
			return 2
		}
		return 0
	}

	i := 0
	for i < len(tokens) {
		t := tokens[i]
		switch t {
		case "+", "-", "*", "/":
			for len(stackOps) > 0 && prec(stackOps[len(stackOps)-1]) >= prec(t) {
				if err := applyOp(); err != nil {
					return nil, err
				}
			}
			stackOps = append(stackOps, t)
		case "(":
			stackOps = append(stackOps, t)
		case ")":
			for len(stackOps) > 0 && stackOps[len(stackOps)-1] != "(" {
				if err := applyOp(); err != nil {
					return nil, err
				}
			}
			if len(stackOps) == 0 {
				return nil, fmt.Errorf("mismatched parentheses")
			}
			stackOps = stackOps[:len(stackOps)-1] // pop "("
		default:
			v, err := strconv.ParseFloat(t, 64)
			if err != nil {
				return nil, err
			}
			stackNums = append(stackNums, v)
		}
		i++
	}

	for len(stackOps) > 0 {
		if stackOps[len(stackOps)-1] == "(" || stackOps[len(stackOps)-1] == ")" {
			return nil, fmt.Errorf("mismatched parentheses")
		}
		if err := applyOp(); err != nil {
			return nil, err
		}
	}

	return stackNums, nil
}

func ensureExt(path, ext string) string {
	if !strings.HasSuffix(strings.ToLower(path), strings.ToLower(ext)) {
		return path + ext
	}
	return path
}

func currentValues(inputs []textinput.Model) []string {
	out := make([]string, len(inputs))
	for i := range inputs {
		out[i] = inputs[i].Value()
	}
	return out
}

func formatMoney(v int) string {
	s := strconv.Itoa(v)
	// insert commas every 3 digits from right
	n := len(s)
	if n <= 3 {
		return s
	}
	var parts []string
	for n > 3 {
		parts = append([]string{s[n-3 : n]}, parts...)
		n -= 3
		if n <= 3 {
			parts = append(parts, s[0:n])
			break
		}
	}
	// reverse parts because we appended in reverse order
	for i, j := 0, len(parts)-1; i < j; i, j = i+1, j-1 {
		parts[i], parts[j] = parts[j], parts[i]
	}
	return strings.Join(parts, ",")
}

func roundTaka(f float64) float64 {
	return math.Round(f)
}

// method on model to open prompt (wrapper)
func (m model) openPrompt(mode promptMode, defaultPath, placeholder string) (model, tea.Cmd) {
	newM := m
	newM.promptMode = mode
	newM.promptActive = true
	newM.prompt.SetValue(defaultPath)
	newM.prompt.Placeholder = placeholder
	newM.state = statePrompt
	return newM, nil
}

func renderTaxPieChartString(tax, surcharge float64, width int) string {
	return renderTaxPieChart(tax, surcharge, width)
}

func formatMoneyFromFloat(f float64) string {
	return formatMoney(int(roundTaka(f)))
}

func renderDetailsPanelString(width int) string {
	return renderDetailsPanel(width)
}

func renderBannerString(w int) string {
	return renderBanner(w)
}

func renderScrollableTextString(content string, offset, height int) string {
	return renderScrollableText(content, offset, height)
}

func main() {
	// Launch TUI
	m := initialModel()
	if err := tea.NewProgram(m, tea.WithAltScreen()).Start(); err != nil {
		fmt.Printf("Error starting program: %v\n", err)
		os.Exit(1)
	}
}

// small helpers used earlier
func maxInt(a, b int) int {
	if a > b {
		return a
	}
	return b
}
func minInt(a, b int) int {
	if a < b {
		return a
	}
	return b
}
