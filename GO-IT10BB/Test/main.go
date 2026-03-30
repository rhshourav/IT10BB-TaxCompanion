package main

import (
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
)

/*
Tax Companion (Bangladesh) - Final single-file TUI
Author: rhshourav  •  github.com/rhshourav

Features:
- Animated ASCII banner with lipgloss gradient sweep
- Details/info panel is responsive and scrollable
- Robust numeric parsing (expressions, 1k, 1 lakh, 1cr, commas)
- Exact expense allocation engine (no taka drift)
- Net-wealth surcharge check and wealth consistency checker
- Pie-style terminal charts
- Salary breakdown / payslip when only gross salary is entered
*/

var (
	titleStyle    = lipgloss.NewStyle().Bold(true).Foreground(lipgloss.Color("#FAFAFA")).Background(lipgloss.Color("#056139")).Padding(0, 1).MarginBottom(1)
	subTitleStyle = lipgloss.NewStyle().Bold(true).Foreground(lipgloss.Color("#056139")).Background(lipgloss.Color("#FAFAFA")).Padding(0, 1).MarginTop(1).MarginBottom(1)
	headerStyle   = lipgloss.NewStyle().Foreground(lipgloss.Color("#00D7FF")).Bold(true)
	authorStyle   = lipgloss.NewStyle().Bold(true).Foreground(lipgloss.Color("#FFFFFF")).Background(lipgloss.Color("#333333")).Padding(0, 1)
	moneyStyle    = lipgloss.NewStyle().Foreground(lipgloss.Color("#00FF00"))
	pctStyle      = lipgloss.NewStyle().Foreground(lipgloss.Color("#FFA500"))
	docStyle      = lipgloss.NewStyle().Margin(1, 2)
	helpStyle     = lipgloss.NewStyle().Foreground(lipgloss.Color("#777777")).MarginTop(1)
	taxStyle      = lipgloss.NewStyle().Foreground(lipgloss.Color("#FF0000")).Bold(true)
	chartLabel    = lipgloss.NewStyle().Bold(true)
	detailKey     = lipgloss.NewStyle().Foreground(lipgloss.Color("#00D7FF")).Bold(true)
	boxStyle      = lipgloss.NewStyle().Border(lipgloss.RoundedBorder()).Padding(1, 1)
)

type appState int

const (
	stateInput appState = iota
	stateLoading
	stateResult
)

type model struct {
	state         appState
	step          int
	inputs        []textinput.Model
	spin          spinner.Model
	authorSpin    spinner.Model
	showInfo      bool
	infoOffset    int
	resultView    string
	width         int
	height        int
	gradientPhase int
	authorLine    string
	authorGithub  string
}

type calculationDoneMsg struct{}

type pieSlice struct {
	Label string
	Value float64
}

func initialModel() model {
	s := spinner.New()
	s.Spinner = spinner.Dot
	s.Style = lipgloss.NewStyle().Foreground(lipgloss.Color("205"))

	a := spinner.New()
	a.Spinner = spinner.Line
	a.Style = lipgloss.NewStyle().Foreground(lipgloss.Color("63"))

	m := model{
		state:         stateInput,
		step:          0,
		spin:          s,
		authorSpin:    a,
		showInfo:      false,
		infoOffset:    0,
		resultView:    "",
		width:         100,
		height:        40,
		gradientPhase: 0,
		authorLine:    "Author: rhshourav",
		authorGithub:  "github.com/rhshourav",
	}

	placeholders := []string{
		"1. Gross Annual Salary (BDT) [Default: 0 / Skip]",
		"2. Festival bonus already included in gross salary? (y/n) [Default: n]",
		"3. Enter Custom Salary Breakdown? (y/n) [Default: n]",
		"   -> Basic Pay (annual BDT) [Default: 0]",
		"   -> House Rent Allowance (annual BDT) [Default: 0]",
		"   -> Medical Allowance (annual BDT) [Default: 0]",
		"   -> Food Allowance (annual BDT) [Default: 0]",
		"   -> Transport / Conveyance (annual BDT) [Default: 0]",
		"   -> Mobile & Other Allowance (annual BDT) [Default: 0]",
		"4. Total Annual Expense (BDT) [Default: 0 / Skip]",
		"5. Location (dhaka/other) [Default: other]",
		"6. Family Size [Default: 3]",
		"7. Do you have kids? (y/n) [Default: n]",
		"8. Do you own your home? (y/n) [Default: n]",
		"9. Home-support staff? (y/n) [Default: n]",
		"10. Mode (balanced/conservative/comfortable) [Default: balanced]",
		"11. Previous Year's Gross Income (BDT) [Default: 0 / Optional]",
		"12. Net Wealth (current) (BDT) [Total assets - liabilities]",
		"13. Opening Net Wealth (previous year) (BDT) [Default: 0]",
		"14. Apply Net Wealth Surcharge? (y/n) [Default: n]",
		"15. Surcharge Percent (number like 10 OR 'auto') [Default: auto]",
	}

	m.inputs = make([]textinput.Model, len(placeholders))
	for i := range m.inputs {
		t := textinput.New()
		t.Placeholder = placeholders[i]
		t.CharLimit = 256
		if i == 0 {
			t.Focus()
		}
		m.inputs[i] = t
	}

	return m
}

func (m model) Init() tea.Cmd {
	return tea.Batch(textinput.Blink)
}

func (m model) Update(msg tea.Msg) (tea.Model, tea.Cmd) {
	switch msg := msg.(type) {
	case tea.KeyMsg:
		switch msg.String() {
		case "ctrl+c", "q", "esc":
			return m, tea.Quit
		case "r":
			if m.state == stateResult {
				return initialModel(), textinput.Blink
			}
		case "i":
			m.showInfo = !m.showInfo
			if m.showInfo {
				m.infoOffset = 0
			}
			return m, nil
		}

		if m.showInfo {
			switch msg.String() {
			case "up", "k":
				if m.infoOffset > 0 {
					m.infoOffset--
				}
				return m, nil
			case "down", "j":
				m.infoOffset++
				return m, nil
			case "pgup":
				m.infoOffset -= (m.height / 2)
				if m.infoOffset < 0 {
					m.infoOffset = 0
				}
				return m, nil
			case "pgdown":
				m.infoOffset += (m.height / 2)
				return m, nil
			case "home":
				m.infoOffset = 0
				return m, nil
			case "end":
				m.infoOffset = 100000
				return m, nil
			}
		}

		if !m.showInfo && m.state == stateInput {
			switch msg.String() {
			case "enter":
				if m.step == 2 {
					wantsCustom := strings.ToLower(getVal(m.inputs[2].Value(), "n")) == "y"
					if !wantsCustom {
						m.step = 9
					} else {
						m.step++
					}
				} else if m.step < len(m.inputs)-1 {
					m.step++
				} else {
					m.state = stateLoading
					return m, tea.Batch(m.spin.Tick, triggerCalculation())
				}

				for i := range m.inputs {
					m.inputs[i].Blur()
				}
				if m.step < len(m.inputs) {
					m.inputs[m.step].Focus()
				}
				return m, nil
			}
		}

	case spinner.TickMsg:
		var cmd tea.Cmd
		m.spin, cmd = m.spin.Update(msg)
		m.authorSpin, _ = m.authorSpin.Update(msg)
		m.gradientPhase = (m.gradientPhase + 1) % 1000000
		return m, cmd

	case calculationDoneMsg:
		m.state = stateResult
		m.resultView = generateResults(m)
		return m, nil

	case tea.WindowSizeMsg:
		m.width = msg.Width
		m.height = msg.Height
		if m.infoOffset < 0 {
			m.infoOffset = 0
		}
	}

	if m.state == stateInput && !m.showInfo {
		var cmd tea.Cmd
		m.inputs[m.step], cmd = m.inputs[m.step].Update(msg)
		return m, cmd
	}

	return m, nil
}

func (m model) View() string {
	author := authorStyle.Render(fmt.Sprintf(" %s %s • %s ", m.authorSpin.View(), m.getAuthorLine(), m.authorGithub))
	banner := renderBanner(m.width, m.gradientPhase, m.authorSpin.View())

	if m.state == stateLoading {
		body := fmt.Sprintf("\n\n   %s Computing tax, surcharge and allocations...\n\n", m.spin.View())
		return docStyle.Render(author + "\n\n" + banner + "\n\n" + body)
	}

	if m.showInfo {
		d := renderDetailsPanel(m.width)
		lines := strings.Split(d, "\n")
		maxBody := m.height - 8
		if maxBody < 6 {
			maxBody = 6
		}
		if m.infoOffset > len(lines)-maxBody {
			m.infoOffset = max(0, len(lines)-maxBody)
		}
		start := m.infoOffset
		end := start + maxBody
		if end > len(lines) {
			end = len(lines)
		}
		visible := strings.Join(lines[start:end], "\n")
		footer := helpStyle.Render("Up/Down/PgUp/PgDn/Home/End → scroll • i → close details • q → quit")
		return docStyle.Render(author + "\n\n" + banner + "\n\n" + visible + "\n\n" + footer)
	}

	if m.state == stateResult {
		footer := helpStyle.Render("r → restart • i → details • q → quit")
		return docStyle.Render(author + "\n\n" + banner + "\n\n" + m.resultView + "\n\n" + footer)
	}

	body := titleStyle.Render(" TAX COMPANION (BANGLADESH) ") + "\n\n"
	for i := 0; i <= m.step && i < len(m.inputs); i++ {
		wantsCustom := strings.ToLower(getVal(m.inputs[2].Value(), "n")) == "y"
		if i >= 3 && i <= 8 && !wantsCustom {
			continue
		}
		body += fmt.Sprintf("%s\n%s\n\n", headerStyle.Render(m.inputs[i].Placeholder), m.inputs[i].View())
	}
	hint := helpStyle.Render("Enter → advance • i → details • q → quit")
	return docStyle.Render(author + "\n\n" + banner + "\n\n" + body + "\n" + hint)
}

func (m model) getAuthorLine() string {
	return m.authorLine
}

func triggerCalculation() tea.Cmd {
	return func() tea.Msg {
		time.Sleep(700 * time.Millisecond)
		return calculationDoneMsg{}
	}
}

var bannerLines = []string{
	"╭────────────────────────────╮",
	"│        rhshourav           │",
	"│   cyber • code • security  │",
	"╰────────────────────────────╯",
}

var gradientColors = []string{
	"#FF6B6B", "#FF8E72", "#FFB56B", "#FFD86B", "#FFF36B",
	"#EAF07A", "#B8F07A", "#7AF0A0", "#5CE0C6", "#5CE0FF",
	"#7AAEFF", "#7B7BFF", "#B27BFF", "#E77BFF", "#FF7BC1",
}

func renderBanner(width int, phase int, spinView string) string {
	out := []string{}
	artWidth := len(stripANSI(bannerLines[0]))

	for r, line := range bannerLines {
		var sb strings.Builder
		for i, ch := range line {
			idx := (i + r*3 + phase/6) % len(gradientColors)
			col := lipgloss.Color(gradientColors[idx])
			sb.WriteString(lipgloss.NewStyle().Foreground(col).Render(string(ch)))
		}
		lineColored := sb.String()
		if width <= artWidth {
			out = append(out, lineColored)
		} else {
			padding := (width - artWidth) / 2
			out = append(out, strings.Repeat(" ", padding)+lineColored)
		}
	}

	if len(out) >= 2 {
		line := out[1]
		sp := " " + spinView
		if len(stripANSI(line))+len(stripANSI(sp)) < width {
			line += sp
		}
		out[1] = line
	}

	return strings.Join(out, "\n")
}

func stripANSI(s string) string {
	re := regexp.MustCompile(`\x1b\[[0-9;]*m`)
	return re.ReplaceAllString(s, "")
}

func renderDetailsPanel(width int) string {
	w := width - 8
	if w < 40 {
		w = 40
	}
	var sb strings.Builder
	sb.WriteString(subTitleStyle.Render(" INPUT DETAILS (what each option does) ") + "\n\n")

	write := func(k, d string) {
		sb.WriteString(detailKey.Render(k) + "\n")
		sb.WriteString(wrapText(d, w) + "\n\n")
	}

	write("Gross Annual Salary", "Total annual salary before adding festival bonus. Accepts arithmetic (e.g., 12809*23) and shorthands (1k, 1 lakh, 1cr).")
	write("Festival bonus already included?", "If 'y', the program will not add an estimated festival bonus. If 'n', it estimates a bonus and adds it to the final total salary.")
	write("Enter Custom Salary Breakdown? (y/n)", "If 'y' you can input exact annual components (basic, hra, medical, food, transport, mobile/other). If 'n' the tool auto-derives a realistic BD split.")
	write("Basic Pay (annual)", "Annual basic salary.")
	write("House Rent Allowance (annual)", "HRA component.")
	write("Medical / Food / Transport / Mobile", "Standard allowances. If left blank and custom breakdown disabled, auto-split will fill them deterministically.")
	write("Total Annual Expense", "Your declared total household expenditure (IT-10BB). The program allocates 100% of this across expense categories with no taka loss.")
	write("Location (dhaka/other)", "Used to adjust accommodation and living weights (Dhaka increases accommodation weight).")
	write("Family Size / Kids / Own home / Home-support staff", "Used to scale allocations (education, food, and home-support).")
	write("Mode (balanced/conservative/comfortable)", "Adjusts discretionary categories like festival/party and home-support.")
	write("Previous Year's Gross Income", "Optional. Used only to compute previous-year tax (for combined views); it is NOT added to current year income.")
	write("Net Wealth (current) & Opening Net Wealth", "Net wealth = total assets − liabilities. Used for surcharge checks and wealth consistency checking.")
	write("Apply Net Wealth Surcharge? / Surcharge Percent", "If chosen and net wealth > Tk 4 Crore, applies a surcharge on tax. 'auto' uses default tiers: 10% / 20% / 35%.")
	note := "Notes: Wealth consistency compares reported wealth increase to a conservative estimate of after-tax savings (gross - expense - tax). Large mismatches suggest undisclosed items or input errors."
	sb.WriteString(wrapText(note, w))
	return boxStyle.Width(w).Render(sb.String())
}

func wrapText(s string, width int) string {
	if width < 20 {
		width = 20
	}
	words := strings.Fields(s)
	var lines []string
	cur := ""
	for _, w := range words {
		if cur == "" {
			cur = w
			continue
		}
		if len(stripANSI(cur))+1+len(w) > width {
			lines = append(lines, cur)
			cur = w
		} else {
			cur += " " + w
		}
	}
	if cur != "" {
		lines = append(lines, cur)
	}
	return strings.Join(lines, "\n")
}

func parseNumeric(s string) float64 {
	s = strings.TrimSpace(s)
	if s == "" {
		return 0
	}
	s = strings.ToLower(strings.ReplaceAll(s, ",", ""))

	unitPatterns := map[*regexp.Regexp]string{
		regexp.MustCompile(`(?i)(\d+(\.\d+)?)\s*(k|thousand)\b`):      "*1000",
		regexp.MustCompile(`(?i)(\d+(\.\d+)?)\s*(m|mn|million)\b`):    "*1000000",
		regexp.MustCompile(`(?i)(\d+(\.\d+)?)\s*(lakh|lac|lacs)\b`):   "*100000",
		regexp.MustCompile(`(?i)(\d+(\.\d+)?)\s*(cr|crore|crores)\b`): "*10000000",
		regexp.MustCompile(`(?i)(\d+(\.\d+)?)\s*(tk|taka)\b`):         "*1",
	}
	for re, repl := range unitPatterns {
		s = re.ReplaceAllString(s, "$1"+repl)
	}

	v, err := evalExpression(s)
	if err != nil {
		if f, e := strconv.ParseFloat(s, 64); e == nil {
			return f
		}
		return 0
	}
	return v
}

func evalExpression(expr string) (float64, error) {
	type token struct {
		typ string
		val string
	}

	tokens := []token{}
	i := 0
	expr = strings.TrimSpace(expr)
	var prev *token

	for i < len(expr) {
		ch := expr[i]
		switch {
		case ch == ' ' || ch == '\t' || ch == '\n':
			i++
		case (ch >= '0' && ch <= '9') || ch == '.':
			j := i
			for j < len(expr) && ((expr[j] >= '0' && expr[j] <= '9') || expr[j] == '.') {
				j++
			}
			tokens = append(tokens, token{typ: "num", val: expr[i:j]})
			prev = &tokens[len(tokens)-1]
			i = j
		case ch == '+' || ch == '-' || ch == '*' || ch == '/':
			if ch == '-' && (prev == nil || prev.typ == "op" || prev.val == "(") {
				tokens = append(tokens, token{typ: "num", val: "0"})
				prev = &tokens[len(tokens)-1]
			}
			tokens = append(tokens, token{typ: "op", val: string(ch)})
			prev = &tokens[len(tokens)-1]
			i++
		case ch == '(' || ch == ')':
			tokens = append(tokens, token{typ: string(ch), val: string(ch)})
			prev = &tokens[len(tokens)-1]
			i++
		default:
			return 0, fmt.Errorf("invalid char: %c", ch)
		}
	}

	out := []token{}
	opStack := []token{}
	prec := map[string]int{"+": 1, "-": 1, "*": 2, "/": 2}

	for _, tk := range tokens {
		if tk.typ == "num" {
			out = append(out, tk)
		} else if tk.typ == "op" {
			for len(opStack) > 0 {
				top := opStack[len(opStack)-1]
				if top.typ == "op" && ((prec[top.val] > prec[tk.val]) || (prec[top.val] == prec[tk.val])) {
					out = append(out, top)
					opStack = opStack[:len(opStack)-1]
					continue
				}
				break
			}
			opStack = append(opStack, tk)
		} else if tk.typ == "(" {
			opStack = append(opStack, tk)
		} else if tk.typ == ")" {
			found := false
			for len(opStack) > 0 {
				top := opStack[len(opStack)-1]
				opStack = opStack[:len(opStack)-1]
				if top.typ == "(" {
					found = true
					break
				}
				out = append(out, top)
			}
			if !found {
				return 0, fmt.Errorf("mismatched parentheses")
			}
		}
	}

	for len(opStack) > 0 {
		top := opStack[len(opStack)-1]
		opStack = opStack[:len(opStack)-1]
		if top.typ == "(" || top.typ == ")" {
			return 0, fmt.Errorf("mismatched parentheses")
		}
		out = append(out, top)
	}

	stack := []float64{}
	for _, tk := range out {
		if tk.typ == "num" {
			f, err := strconv.ParseFloat(tk.val, 64)
			if err != nil {
				return 0, err
			}
			stack = append(stack, f)
		} else if tk.typ == "op" {
			if len(stack) < 2 {
				return 0, fmt.Errorf("invalid expression")
			}
			b := stack[len(stack)-1]
			a := stack[len(stack)-2]
			stack = stack[:len(stack)-2]
			var r float64
			switch tk.val {
			case "+":
				r = a + b
			case "-":
				r = a - b
			case "*":
				r = a * b
			case "/":
				if b == 0 {
					return 0, fmt.Errorf("division by zero")
				}
				r = a / b
			default:
				return 0, fmt.Errorf("unknown op")
			}
			stack = append(stack, r)
		}
	}

	if len(stack) != 1 {
		return 0, fmt.Errorf("invalid final stack")
	}
	return stack[0], nil
}

func calculateTax(taxable float64) (string, float64, map[string]float64) {
	out := make(map[string]float64)
	if taxable <= 350000 {
		out["total_tax"] = 0
		return "🎉 No tax liability! (Income at or below Tk 3,50,000)\n", 0, out
	}

	slabs := []struct {
		limit float64
		rate  float64
		label string
	}{
		{350000, 0.00, "First 3.5 Lakh (0%)"},
		{100000, 0.05, "Next 1 Lakh (5%)"},
		{400000, 0.10, "Next 4 Lakh (10%)"},
		{500000, 0.15, "Next 5 Lakh (15%)"},
		{500000, 0.20, "Next 5 Lakh (20%)"},
		{math.MaxFloat64, 0.25, "Remaining (25%)"},
	}

	var sb strings.Builder
	var totalTax float64
	remaining := taxable

	sb.WriteString(fmt.Sprintf("%-22s | %-16s | %s\n", "SLAB", "AMOUNT TAXED", "TAX GENERATED"))
	sb.WriteString(strings.Repeat("-", 62) + "\n")

	for _, slab := range slabs {
		if remaining <= 0 {
			break
		}
		amt := math.Min(remaining, slab.limit)
		t := amt * slab.rate
		totalTax += t
		remaining -= amt

		amtStr := pctStyle.Render(fmt.Sprintf("Tk %10s", formatMoney(int(amt))))
		taxStr := moneyStyle.Render(fmt.Sprintf("Tk %10s", formatMoney(int(t))))
		if t > 0 {
			taxStr = taxStyle.Render(fmt.Sprintf("Tk %10s", formatMoney(int(t))))
		}
		sb.WriteString(fmt.Sprintf("%-22s | %16s | %16s\n", slab.label, amtStr, taxStr))
		out[slab.label] = t
	}

	sb.WriteString(strings.Repeat("-", 62) + "\n")
	if totalTax > 0 && totalTax < 3000 {
		totalTax = 3000
		sb.WriteString("* Applied minimum tax threshold (Tk 3,000)\n")
	}
	sb.WriteString(fmt.Sprintf("%-22s | %-16s | %s\n", "TOTAL TAX PAYABLE", "", taxStyle.Render(fmt.Sprintf("Tk %10s", formatMoney(int(totalTax))))))
	out["total_tax"] = totalTax
	return sb.String(), totalTax, out
}

func deriveSalaryBreakdownFromGross(gross float64) (int, int, int, int) {
	if gross <= 0 {
		return 0, 0, 0, 0
	}

	// Matches the style of the example:
	// Basic 58%, House rent 29%, Medical 10%, Conveyance 3%
	basic := int(math.Round(gross * 0.58))
	hra := int(math.Round(gross * 0.29))
	med := int(math.Round(gross * 0.10))
	conv := int(math.Round(gross * 0.03))

	sum := basic + hra + med + conv
	diff := int(math.Round(gross)) - sum
	if diff != 0 {
		conv += diff
	}
	if conv < 0 {
		conv = 0
	}
	return basic, hra, med, conv
}

func estimateFestivalBonus(gross float64) int {
	if gross <= 0 {
		return 0
	}
	// Estimate used when bonus is not already included.
	// Kept close to the style of the example.
	return int(math.Round(gross * 0.124))
}

func generateResults(m model) string {
	var sb strings.Builder

	gross := parseNumeric(getVal(m.inputs[0].Value(), "0"))
	bonusIncluded := strings.ToLower(getVal(m.inputs[1].Value(), "n")) == "y"
	wantsCustom := strings.ToLower(getVal(m.inputs[2].Value(), "n")) == "y"

	cBasic := parseNumeric(getVal(m.inputs[3].Value(), "0"))
	cHra := parseNumeric(getVal(m.inputs[4].Value(), "0"))
	cMed := parseNumeric(getVal(m.inputs[5].Value(), "0"))
	cFood := parseNumeric(getVal(m.inputs[6].Value(), "0"))
	cTrans := parseNumeric(getVal(m.inputs[7].Value(), "0"))
	cMob := parseNumeric(getVal(m.inputs[8].Value(), "0"))

	totalExpense := parseNumeric(getVal(m.inputs[9].Value(), "0"))
	loc := strings.ToLower(getVal(m.inputs[10].Value(), "other"))
	fam, _ := strconv.Atoi(getVal(m.inputs[11].Value(), "3"))
	kids := strings.ToLower(getVal(m.inputs[12].Value(), "n")) == "y"
	own := strings.ToLower(getVal(m.inputs[13].Value(), "n")) == "y"
	staff := strings.ToLower(getVal(m.inputs[14].Value(), "n")) == "y"
	mode := strings.ToLower(getVal(m.inputs[15].Value(), "balanced"))
	prevIncome := parseNumeric(getVal(m.inputs[16].Value(), "0"))
	netWealthCurrent := parseNumeric(getVal(m.inputs[17].Value(), "0"))
	openingNetWealth := parseNumeric(getVal(m.inputs[18].Value(), "0"))
	applySurcharge := strings.ToLower(getVal(m.inputs[19].Value(), "n")) == "y"
	surchargeInput := strings.ToLower(getVal(m.inputs[20].Value(), "auto"))

	var basic, hra, med, conv int
	var festivalBonus int
	var totalSalary int

	if gross > 0 || wantsCustom {
		sb.WriteString(subTitleStyle.Render(" SALARY BREAKDOWN & TAXABLE INCOME ") + "\n")

		if wantsCustom {
			basic = int(math.Round(cBasic))
			hra = int(math.Round(cHra))
			med = int(math.Round(cMed))
			conv = int(math.Round(cTrans))
			festivalBonus = int(math.Round(float64(basic) / 6.0))

			sum := basic + hra + med + conv + festivalBonus + int(math.Round(cFood)) + int(math.Round(cMob))
			if sum != int(math.Round(gross)) {
				// Keep custom input usable, but make the sum exact.
				adjust := int(math.Round(gross)) - sum
				conv += adjust
				if conv < 0 {
					conv = 0
				}
			}

			totalSalary = int(math.Round(gross))
		} else {
			basic, hra, med, conv = deriveSalaryBreakdownFromGross(gross)
			if bonusIncluded {
				festivalBonus = 0
				totalSalary = int(math.Round(gross))
			} else {
				festivalBonus = estimateFestivalBonus(gross)
				totalSalary = int(math.Round(gross)) + festivalBonus
			}
		}

		exempt := math.Min(gross/3.0, 450000)
		taxable := gross - exempt

		sb.WriteString(fmt.Sprintf("%-30s | %s\n", "Basic Salary", moneyStyle.Render(fmt.Sprintf("Tk %10s", formatMoney(basic)))))
		sb.WriteString(fmt.Sprintf("%-30s | %s\n", "House rent", moneyStyle.Render(fmt.Sprintf("Tk %10s", formatMoney(hra)))))
		sb.WriteString(fmt.Sprintf("%-30s | %s\n", "Medical Allowance", moneyStyle.Render(fmt.Sprintf("Tk %10s", formatMoney(med)))))
		sb.WriteString(fmt.Sprintf("%-30s | %s\n", "Conveyance Allowance", moneyStyle.Render(fmt.Sprintf("Tk %10s", formatMoney(conv)))))
		sb.WriteString(fmt.Sprintf("%-30s | %s\n", "Gross salary", moneyStyle.Render(fmt.Sprintf("Tk %10s", formatMoney(int(math.Round(gross)))))))
		sb.WriteString(fmt.Sprintf("%-30s | %s\n", "festival Bonus", moneyStyle.Render(fmt.Sprintf("Tk %10s", formatMoney(festivalBonus)))))
		sb.WriteString(strings.Repeat("-", 60) + "\n")
		sb.WriteString(fmt.Sprintf("%-30s | %s\n", "Total Salary", moneyStyle.Bold(true).Render(fmt.Sprintf("Tk %10s", formatMoney(totalSalary)))))
		sb.WriteString(strings.Repeat("-", 60) + "\n")
		sb.WriteString(fmt.Sprintf("%-30s | %s\n\n", "NET TAXABLE INCOME", taxStyle.Render(fmt.Sprintf("Tk %10s", formatMoney(int(math.Round(taxable)))))))

		if !wantsCustom {
			sb.WriteString(subTitleStyle.Render(" BREAKDOWN OF TOTAL SALARY ") + "\n")
			sb.WriteString(fmt.Sprintf("%-30s | %s\n", "Basic Salary", moneyStyle.Render(fmt.Sprintf("Tk %10s", formatMoney(basic)))))
			sb.WriteString(fmt.Sprintf("%-30s | %s\n", "House rent", moneyStyle.Render(fmt.Sprintf("Tk %10s", formatMoney(hra)))))
			sb.WriteString(fmt.Sprintf("%-30s | %s\n", "Medical Allowance", moneyStyle.Render(fmt.Sprintf("Tk %10s", formatMoney(med)))))
			sb.WriteString(fmt.Sprintf("%-30s | %s\n", "Conveyance Allowance", moneyStyle.Render(fmt.Sprintf("Tk %10s", formatMoney(conv)))))
			sb.WriteString(fmt.Sprintf("%-30s | %s\n", "Gross salary", moneyStyle.Render(fmt.Sprintf("Tk %10s", formatMoney(int(math.Round(gross)))))))
			sb.WriteString(fmt.Sprintf("%-30s | %s\n", "festival Bonus", moneyStyle.Render(fmt.Sprintf("Tk %10s", formatMoney(festivalBonus)))))
			sb.WriteString(fmt.Sprintf("%-30s | %s\n", "Total Salary", moneyStyle.Bold(true).Render(fmt.Sprintf("Tk %10s", formatMoney(totalSalary)))))
			sb.WriteString("\n")
		}

		sb.WriteString(subTitleStyle.Render(" TAX LIABILITY (CURRENT YEAR - SLAB BREAKDOWN) ") + "\n")
		taxText, totalTaxCur, slabMap := calculateTax(taxable)
		sb.WriteString(taxText + "\n")

		var totalTaxPrev float64
		if prevIncome > 0 {
			sb.WriteString(subTitleStyle.Render(" PREVIOUS YEAR TAX (OPTIONAL) ") + "\n")
			prevExempt := math.Min(prevIncome/3.0, 450000)
			prevTaxable := prevIncome - prevExempt
			prevTaxText, prevTax, _ := calculateTax(prevTaxable)
			totalTaxPrev = prevTax
			sb.WriteString(fmt.Sprintf("%-30s | %s\n", "PREVIOUS GROSS INCOME", moneyStyle.Render(fmt.Sprintf("Tk %10s", formatMoney(int(math.Round(prevIncome)))))))
			sb.WriteString(fmt.Sprintf("%-30s | %s\n", "PREVIOUS EXEMPTION", pctStyle.Render(fmt.Sprintf("Tk %10s", formatMoney(int(math.Round(prevExempt)))))))
			sb.WriteString(fmt.Sprintf("%-30s | %s\n\n", "PREVIOUS TAXABLE INCOME", taxStyle.Render(fmt.Sprintf("Tk %10s", formatMoney(int(math.Round(prevTaxable)))))))
			sb.WriteString(prevTaxText + "\n")
		}

		combinedBeforeSurcharge := totalTaxCur + totalTaxPrev
		sb.WriteString(strings.Repeat("=", 60) + "\n")
		sb.WriteString(fmt.Sprintf("%-30s | %s\n", "TOTAL TAX (CURRENT)", taxStyle.Render(fmt.Sprintf("Tk %10s", formatMoney(int(math.Round(totalTaxCur)))))))
		if prevIncome > 0 {
			sb.WriteString(fmt.Sprintf("%-30s | %s\n", "TOTAL TAX (PREVIOUS)", taxStyle.Render(fmt.Sprintf("Tk %10s", formatMoney(int(math.Round(totalTaxPrev)))))))
		}
		sb.WriteString(fmt.Sprintf("%-30s | %s\n\n", "COMBINED (BEFORE SURCHARGE)", taxStyle.Render(fmt.Sprintf("Tk %10s", formatMoney(int(math.Round(combinedBeforeSurcharge)))))))

		var surchargeRate float64
		var surchargeAmount float64
		threshold := 40000000.0

		if applySurcharge && netWealthCurrent > threshold {
			if surchargeInput == "auto" || surchargeInput == "" {
				switch {
				case netWealthCurrent <= 50000000:
					surchargeRate = 0.10
				case netWealthCurrent <= 100000000:
					surchargeRate = 0.20
				default:
					surchargeRate = 0.35
				}
			} else {
				pStr := strings.TrimSuffix(surchargeInput, "%")
				if p, err := strconv.ParseFloat(pStr, 64); err == nil {
					surchargeRate = p / 100.0
				} else {
					surchargeRate = 0.10
				}
			}
			surchargeAmount = combinedBeforeSurcharge * surchargeRate
			sb.WriteString(subTitleStyle.Render(" NET WEALTH SURCHARGE CHECK ") + "\n")
			sb.WriteString(fmt.Sprintf("%-30s | %s\n", "NET WEALTH (user)", moneyStyle.Render(fmt.Sprintf("Tk %10s", formatMoney(int(math.Round(netWealthCurrent)))))))
			sb.WriteString(fmt.Sprintf("%-30s | %s\n", "THRESHOLD", pctStyle.Render("Tk 40,000,000")))
			sb.WriteString(fmt.Sprintf("%-30s | %s\n", "SURCHARGE %", pctStyle.Render(fmt.Sprintf("%.1f%%", surchargeRate*100))))
			sb.WriteString(fmt.Sprintf("%-30s | %s\n\n", "SURCHARGE AMOUNT", taxStyle.Render(fmt.Sprintf("Tk %10s", formatMoney(int(math.Round(surchargeAmount)))))))
		} else {
			if applySurcharge {
				sb.WriteString(subTitleStyle.Render(" NET WEALTH SURCHARGE CHECK ") + "\n")
				sb.WriteString(fmt.Sprintf("%-30s | %s\n\n", "RESULT", pctStyle.Render("No surcharge applied (below threshold)")))
			}
		}

		finalTax := combinedBeforeSurcharge + surchargeAmount
		sb.WriteString(strings.Repeat("=", 60) + "\n")
		sb.WriteString(fmt.Sprintf("%-30s | %s\n", "FINAL TAX (AFTER SURCHARGE)", taxStyle.Render(fmt.Sprintf("Tk %10s", formatMoney(int(math.Round(finalTax)))))))
		sb.WriteString(strings.Repeat("=", 60) + "\n\n")

		sb.WriteString(subTitleStyle.Render(" VISUAL: TAX SUMMARY (PIE CHART) ") + "\n")
		sb.WriteString(renderTaxPieChart(totalTaxCur, surchargeAmount, 56) + "\n\n")
		sb.WriteString(subTitleStyle.Render(" TAX SLAB CONTRIBUTIONS (PIE CHART) ") + "\n")
		sb.WriteString(renderSlabPieChart(slabMap, 56) + "\n\n")

		sb.WriteString(subTitleStyle.Render(" WEALTH CONSISTENCY CHECK ") + "\n")
		wealthIncrease := netWealthCurrent - openingNetWealth
		expectedSavings := gross - totalExpense - finalTax
		sb.WriteString(fmt.Sprintf("%-40s : %s\n", "Current Net Wealth (reported)", moneyStyle.Render(fmt.Sprintf("Tk %s", formatMoney(int(math.Round(netWealthCurrent)))))))
		sb.WriteString(fmt.Sprintf("%-40s : %s\n", "Opening Net Wealth (previous year)", moneyStyle.Render(fmt.Sprintf("Tk %s", formatMoney(int(math.Round(openingNetWealth)))))))
		sb.WriteString(fmt.Sprintf("%-40s : %s\n", "Wealth Increase (current - opening)", moneyStyle.Render(fmt.Sprintf("Tk %s", formatMoney(int(math.Round(wealthIncrease)))))))
		sb.WriteString(fmt.Sprintf("%-40s : %s\n\n", "Estimated After-tax Savings (gross - expense - tax)", moneyStyle.Render(fmt.Sprintf("Tk %s", formatMoney(int(math.Round(expectedSavings)))))))

		tol := math.Max(10000.0, math.Abs(expectedSavings)*0.01)
		diff := wealthIncrease - expectedSavings
		if math.Abs(diff) <= tol {
			sb.WriteString(fmt.Sprintf("Result: %s\n\n", pctStyle.Render("OK — wealth increase matches after-tax savings within tolerance")))
		} else if diff > 0 {
			sb.WriteString(fmt.Sprintf("Result: %s\n", taxStyle.Render("ALERT — unexplained wealth increase")))
			sb.WriteString(fmt.Sprintf("Details: Reported wealth rose by %s more than estimated savings (difference = %s).\n", moneyStyle.Render(fmt.Sprintf("Tk %s", formatMoney(int(math.Round(diff))))), moneyStyle.Render(fmt.Sprintf("Tk %s", formatMoney(int(math.Round(diff)))))))
			sb.WriteString("Possible reasons: undisclosed income, asset revaluation, gifts, inheritances, loans forgiven, or incorrect inputs.\n")
			sb.WriteString("Suggested actions: verify asset/liability entries, check bank balances, add explanations for large transfers.\n\n")
		} else {
			sb.WriteString(fmt.Sprintf("Result: %s\n", pctStyle.Render("Wealth increase is LESS than estimated savings")))
			sb.WriteString("Possible reasons: cash withdrawn, spending not entered as expense, loan repayments, or data-entry mismatch.\n\n")
		}

		_ = slabMap
	}

	if totalExpense > 0 {
		sb.WriteString(subTitleStyle.Render(" IT-10BB FORM ESTIMATE RESULTS ") + "\n")
		pcts, amts := computeAllocation(totalExpense, loc, fam, kids, own, staff, mode)

		expKeys := []string{
			"Food, Clothing and Essentials",
			"Accommodation Expense",
			"Electricity",
			"Gas, Water, Sewer and Garbage",
			"Phone, Internet, TV & Subs",
			"Home-Support & Other Expenses",
			"Education Expenses",
			"Festival, Party, Events",
		}

		for _, k := range expKeys {
			if pcts[k] == 0 && amts[k] == 0 {
				continue
			}
			pStr := pctStyle.Render(fmt.Sprintf("%3d%%", pcts[k]))
			aStr := moneyStyle.Render(fmt.Sprintf("Tk %10s", formatMoney(amts[k])))
			sb.WriteString(fmt.Sprintf("%-30s | %s | %s\n", k, pStr, aStr))
		}

		sb.WriteString(strings.Repeat("-", 55) + "\n")
		totalStr := moneyStyle.Bold(true).Render(fmt.Sprintf("Tk %10s", formatMoney(int(math.Round(totalExpense)))))
		sb.WriteString(fmt.Sprintf("%-30s | %s | %s\n", "TOTAL ANNUAL EXPENDITURE", pctStyle.Render("100%"), totalStr))
		sb.WriteString("\n" + subTitleStyle.Render(" VISUAL: EXPENSE ALLOCATION (PIE CHART) ") + "\n")
		sb.WriteString(renderExpensePieChart(pcts, amts, 56) + "\n")
	} else if gross == 0 {
		sb.WriteString("No income or expense provided. Nothing to calculate!\n")
	}

	return sb.String()
}

func renderPieChart(title string, items []pieSlice, width int) string {
	filtered := make([]pieSlice, 0, len(items))
	for _, it := range items {
		if it.Value > 0 {
			filtered = append(filtered, it)
		}
	}
	if len(filtered) == 0 {
		return "No data to display."
	}

	sort.Slice(filtered, func(i, j int) bool { return filtered[i].Value > filtered[j].Value })

	total := 0.0
	for _, it := range filtered {
		total += it.Value
	}
	if total <= 0 {
		return "No data to display."
	}

	palette := []string{"#00D7FF", "#7CFF6B", "#FFD166", "#FF6B6B", "#A78BFA", "#F472B6", "#34D399", "#F59E0B", "#60A5FA", "#F97316"}
	segmentCount := 24
	segs := make([]string, segmentCount)

	for i := 0; i < segmentCount; i++ {
		pos := (float64(i) + 0.5) / float64(segmentCount)
		cumulative := 0.0
		picked := len(filtered) - 1
		for idx, it := range filtered {
			share := it.Value / total
			if pos <= cumulative+share || idx == len(filtered)-1 {
				picked = idx
				break
			}
			cumulative += share
		}
		segs[i] = lipgloss.NewStyle().Foreground(lipgloss.Color(palette[picked%len(palette)])).Render("●")
	}

	ring := strings.Join(segs, "")
	var sb strings.Builder
	sb.WriteString(title + "\n")
	if width > len(stripANSI(ring)) {
		sb.WriteString(strings.Repeat(" ", (width-len(stripANSI(ring)))/2))
	}
	sb.WriteString(ring + "\n\n")

	for i, it := range filtered {
		pct := (it.Value / total) * 100
		bullet := lipgloss.NewStyle().Foreground(lipgloss.Color(palette[i%len(palette)])).Render("●")
		sb.WriteString(fmt.Sprintf("%s %-28s | %3.0f%% | Tk %s\n", bullet, it.Label, pct, formatMoney(int(math.Round(it.Value)))))
	}

	return sb.String()
}

func renderTaxPieChart(tax, surcharge float64, width int) string {
	items := []pieSlice{{Label: "Current Tax", Value: tax}}
	if surcharge > 0 {
		items = append(items, pieSlice{Label: "Surcharge", Value: surcharge})
	}
	return renderPieChart("Tax Composition", items, width)
}

func renderSlabPieChart(slab map[string]float64, width int) string {
	type kv struct {
		k string
		v float64
	}
	arr := []kv{}
	for k, v := range slab {
		if k == "total_tax" || v <= 0 {
			continue
		}
		arr = append(arr, kv{k, v})
	}
	if len(arr) == 0 {
		return "No slab contributions to display."
	}
	sort.Slice(arr, func(i, j int) bool { return arr[i].v > arr[j].v })
	if len(arr) > 8 {
		arr = arr[:8]
	}
	items := make([]pieSlice, 0, len(arr))
	for _, x := range arr {
		items = append(items, pieSlice{Label: x.k, Value: x.v})
	}
	return renderPieChart("Tax Slab Contributions", items, width)
}

func renderExpensePieChart(pcts map[string]int, amts map[string]int, width int) string {
	type kv struct {
		k string
		v float64
	}
	arr := []kv{}
	for k, p := range pcts {
		if p <= 0 {
			continue
		}
		arr = append(arr, kv{k, float64(amts[k])})
	}
	if len(arr) == 0 {
		return "No expenses to chart."
	}
	sort.Slice(arr, func(i, j int) bool { return arr[i].v > arr[j].v })
	if len(arr) > 8 {
		arr = arr[:8]
	}
	items := make([]pieSlice, 0, len(arr))
	for _, x := range arr {
		items = append(items, pieSlice{Label: x.k, Value: x.v})
	}
	return renderPieChart("Expense Allocation", items, width)
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
		weights["Education Expenses"] = 0.0
	}
	if loc == "dhaka" || loc == "city" || loc == "metro" {
		weights["Accommodation Expense"] *= 1.20
		weights["Food, Clothing and Essentials"] *= 1.05
		weights["Home-Support & Other Expenses"] *= 1.10
	} else {
		weights["Accommodation Expense"] *= 0.90
	}

	extra := math.Max(0, float64(familySize-2))
	if extra > 0 {
		weights["Food, Clothing and Essentials"] *= (1 + 0.05*extra)
		if hasKids {
			weights["Education Expenses"] *= (1 + 0.04*extra)
		}
	}

	if ownHome {
		weights["Accommodation Expense"] *= 0.60
		weights["Home-Support & Other Expenses"] *= 1.05
	}
	if !staff {
		weights["Home-Support & Other Expenses"] *= 0.4
	}

	switch mode {
	case "conservative":
		weights["Festival, Party, Events"] *= 0.6
		weights["Home-Support & Other Expenses"] *= 0.7
		weights["Food, Clothing and Essentials"] *= 1.08
	case "comfortable":
		weights["Festival, Party, Events"] *= 1.3
		weights["Home-Support & Other Expenses"] *= 1.2
		weights["Food, Clothing and Essentials"] *= 1.05
	}

	totalWeight := 0.0
	for _, v := range weights {
		totalWeight += v
	}

	intPcts := map[string]int{}
	remainders := []struct {
		k string
		r float64
	}{}
	sumPct := 0

	for k, v := range weights {
		raw := (v / totalWeight) * 100.0
		base := int(raw)
		intPcts[k] = base
		sumPct += base
		remainders = append(remainders, struct {
			k string
			r float64
		}{k, raw - float64(base)})
	}

	sort.Slice(remainders, func(i, j int) bool { return remainders[i].r > remainders[j].r })
	for i := 0; i < (100 - sumPct); i++ {
		intPcts[remainders[i%len(remainders)].k]++
	}

	intAmts := map[string]int{}
	amtRem := []struct {
		k string
		r float64
	}{}
	totalInt := int(math.Round(total))
	allocated := 0

	for k, p := range intPcts {
		rawAmt := float64(totalInt) * float64(p) / 100.0
		base := int(rawAmt)
		intAmts[k] = base
		allocated += base
		amtRem = append(amtRem, struct {
			k string
			r float64
		}{k, rawAmt - float64(base)})
	}

	drift := totalInt - allocated
	if drift > 0 {
		sort.Slice(amtRem, func(i, j int) bool { return amtRem[i].r > amtRem[j].r })
		for i := 0; i < drift; i++ {
			intAmts[amtRem[i%len(amtRem)].k]++
		}
	} else if drift < 0 {
		sort.Slice(amtRem, func(i, j int) bool { return amtRem[i].r < amtRem[j].r })
		for i := 0; i < -drift; i++ {
			intAmts[amtRem[i%len(amtRem)].k]--
		}
	}

	return intPcts, intAmts
}

func getVal(inp, def string) string {
	t := strings.TrimSpace(inp)
	if t == "" {
		return def
	}
	return t
}

func formatMoney(n int) string {
	neg := n < 0
	if neg {
		n = -n
	}
	s := strconv.Itoa(n)
	if len(s) <= 3 {
		if neg {
			return "-" + s
		}
		return s
	}
	parts := []string{}
	for i := len(s); i > 0; i -= 3 {
		start := i - 3
		if start < 0 {
			start = 0
		}
		parts = append([]string{s[start:i]}, parts...)
	}
	out := strings.Join(parts, ",")
	if neg {
		return "-" + out
	}
	return out
}

func max(a, b int) int {
	if a > b {
		return a
	}
	return b
}

func main() {
	p := tea.NewProgram(initialModel(), tea.WithAltScreen())
	if _, err := p.Run(); err != nil {
		fmt.Printf("Error: %v\n", err)
		os.Exit(1)
	}
}
