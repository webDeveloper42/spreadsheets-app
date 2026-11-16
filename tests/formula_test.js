// The formula engine against answers Excel and Google Sheets give. Run: node tests/formula_test.js [path to formula.js]
"use strict";
const path = require("path");
const F = require(process.argv[2] || path.join(process.env.HOME, "capture-app/capture_app/static/sheets/formula.js"));
let fails = 0, passes = 0;
function book(cells, more = []) {
  const doc = { sheets: [{ name: "Sheet1", rows: 100, cols: 26, cells: {} }, ...more.map((m) => ({ rows: 100, cols: 26, ...m }))] };
  for (const [k, v] of Object.entries(cells)) doc.sheets[0].cells[k] = { v: F.parseInput(v).v };
  const b = new F.Book(doc); b.recalc(); return b;
}
function val(b, a, si = 0) { const p = F.parseAddr(a); return b.value(si, p.r, p.c); }
function show(v) { return F.isErr(v) ? v.code : JSON.stringify(v); }
function eq(got, want, what) {
  let ok;
  if (F.isErr(got)) ok = typeof want === "string" && want.startsWith("#") && got.code === want;
  else if (typeof want === "number" && typeof got === "number") ok = Math.abs(got - want) <= 1e-9 * Math.max(1, Math.abs(want));
  else ok = got === want;
  if (ok) passes++; else { fails++; console.log(`FAIL ${what}: got ${show(got)}, want ${show(want)}`); }
}
// =formula on a small sheet, compared with the expected value
const DATA = { A1: "10", A2: "20", A3: "30", A4: "", A5: "hello", B1: "apple", B2: "banana", B3: "apple", B4: "cherry", B5: "TRUE", C1: "1", C2: "2", C3: "3", C4: "4", C5: "5",
  D1: "Name", E1: "Score", D2: "Ann", E2: "88", D3: "Bob", E3: "95", D4: "Cid", E4: "72", F1: "2024-03-15", F2: "2024-01-31", G1: "$1,200.50", G2: "15%" };
function f(formula, want, extra = {}) {
  const b = book({ ...DATA, ...extra, Z1: formula });
  eq(val(b, "Z1"), want, formula);
  return b;
}

// ---- input parsing
eq(F.parseInput("1,234.5").v, 1234.5, "parse 1,234.5");
eq(F.parseInput("$12").fmt, "$#,##0", "parse $12 format");
eq(F.parseInput("50%").v, 0.5, "parse 50%");
eq(F.parseInput("-$3.25").v, -3.25, "parse -$3.25");
eq(F.parseInput("2024-03-15").v, 45366, "parse ISO date");
eq(F.parseInput("3/15/2024").v, 45366, "parse US date");
eq(F.parseInput("Mar 15, 2024").v, 45366, "parse Mar 15, 2024");
eq(F.parseInput("2/30/2024").v, "2/30/2024", "31 Feb stays text");
eq(F.parseInput("14:30").v, 14.5 / 24, "parse 14:30");
eq(F.parseInput("2:30 pm").v, 14.5 / 24, "parse 2:30 pm");
eq(F.parseInput("true").v, true, "parse true");
eq(F.parseInput("'007").v, "'007", "forced text");
eq(F.parseInput("1e3").v, 1000, "parse 1e3");
eq(F.parseInput("12 apples").v, "12 apples", "text with a number");

// ---- formats
const fm = (v, c) => F.format(v, c).text;
eq(fm(1234.5, "#,##0.00"), "1,234.50", "fmt #,##0.00");
eq(fm(1234.5, "$#,##0.00"), "$1,234.50", "fmt money");
eq(fm(-5, "$#,##0.00"), "-$5.00", "fmt negative money");
eq(fm(-5, "$#,##0.00;($#,##0.00)"), "($5.00)", "fmt negative in brackets");
eq(fm(0.256, "0%"), "26%", "fmt percent");
eq(fm(0.256, "0.00%"), "25.60%", "fmt percent 2dp");
eq(fm(12345, "0.00E+00"), "1.23E+04", "fmt scientific");
eq(fm(7, "000"), "007", "fmt leading zeros");
eq(fm(3.14159, "0.00"), "3.14", "fmt 2dp");
eq(fm(2.5, "0"), "3", "fmt round half up");
eq(fm(1.005, "0.00"), "1.01", "fmt float rounding");
eq(fm(0.5, "#.##"), ".5", "fmt #.##");
eq(fm(45366, "yyyy-mm-dd"), "2024-03-15", "fmt iso date");
eq(fm(45366, "m/d/yyyy"), "3/15/2024", "fmt us date");
eq(fm(45366, "mmm d, yyyy"), "Mar 15, 2024", "fmt mmm d");
eq(fm(45366, "dddd"), "Friday", "fmt weekday");
eq(fm(45366.75, "h:mm AM/PM"), "6:00 PM", "fmt time pm");
eq(fm(45366.5 + 1 / 1440 * 5, "hh:mm"), "12:05", "fmt hh:mm minutes not months");
eq(fm(1.5, "[h]:mm"), "36:00", "fmt duration");
eq(fm(1234567, "#,##0,\"K\""), "1,235K", "fmt thousands scaling");
eq(fm(0.1 + 0.2, undefined), "0.3", "general 0.1+0.2");
eq(fm(1 / 3, undefined), "0.33333333333", "general 1/3");
eq(fm(123456789012, undefined), "1.23457E+11", "general big");
eq(F.format(-1, "0;[Red]-0").color, "red", "fmt red section");
eq(fm("hi", "@"), "hi", "text format");
eq(fm(0, "0.00;-0.00;\"zero\""), "zero", "zero section");

// ---- operators
f("=1+2*3", 7); f("=(1+2)*3", 9); f("=-2^2", 4); f("=2^3^2", 64); f("=10%", 0.1); f("=50%*A1", 5);
f('="a"&"b"&1', "ab1"); f("=1/0", "#DIV/0!"); f('="3"+1', 4); f('="x"+1', "#VALUE!"); f("=TRUE+1", 2);
f("=A4+1", 1); f('=A4&"x"', "x"); f("=A1=10", true); f('="A"="a"', true); f("=1<\"a\"", true); f('="b">"a"', true);
f("=1<>2", true); f("=A1:A3*2", 20); f("=SUM(A1:A3*C1:C3)", 140); f("=0.1+0.2=0.3", true);
f("=#N/A+1", "#N/A"); f("=SUM(", "#ERROR!"); f("=FOO(1)", "#NAME?"); f("=foo", "#NAME?"); f("=Sheet9!A1", "#REF!");

// ---- math
f("=SUM(A1:A5)", 60); f("=SUM(A1:A3,5,\"5\")", 70); f("=SUM(B1:B5)", 0); f("=PRODUCT(C1:C5)", 120); f("=SUMSQ(C1:C3)", 14);
f('=SUMIF(B1:B4,"apple",C1:C4)', 4); f('=SUMIF(A1:A3,">15")', 50); f('=SUMIF(B1:B4,"*an*",C1:C4)', 2);
f('=SUMIFS(C1:C4,B1:B4,"apple",A1:A4,">15")', 3); f("=SUMPRODUCT(A1:A3,C1:C3)", 140); f("=SUMPRODUCT((B1:B4=\"apple\")*C1:C4)", 4);
f("=ROUND(2.5,0)", 3); f("=ROUND(-2.5,0)", -3); f("=ROUND(1234.567,-2)", 1200); f("=ROUND(1.005,2)", 1.01); f("=ROUNDUP(3.2,0)", 4); f("=ROUNDDOWN(-3.7,0)", -3);
f("=INT(-3.5)", -4); f("=TRUNC(-3.5)", -3); f("=MOD(-3,2)", 1); f("=MOD(3,-2)", -1); f("=MOD(1,0)", "#DIV/0!"); f("=QUOTIENT(7,2)", 3);
f("=POWER(2,10)", 1024); f("=SQRT(16)", 4); f("=SQRT(-1)", "#NUM!"); f("=ABS(-3)", 3); f("=SIGN(-2)", -1); f("=LN(EXP(1))", 1); f("=LOG(100)", 2); f("=LOG(8,2)", 3);
f("=CEILING(4.1,2)", 6); f("=FLOOR(4.9,2)", 4); f("=MROUND(7,3)", 6); f("=EVEN(3)", 4); f("=ODD(2)", 3); f("=FACT(5)", 120); f("=COMBIN(5,2)", 10); f("=PERMUT(5,2)", 20);
f("=GCD(12,18)", 6); f("=LCM(4,6)", 12); f("=PI()", Math.PI); f("=SUBTOTAL(9,A1:A3)", 60); f("=SUBTOTAL(101,A1:A3)", 20); f("=ROMAN(1994)", "MCMXCIV"); f("=BASE(10,2)", "1010"); f("=DECIMAL(\"FF\",16)", 255);
f("=SUM(SEQUENCE(4))", 10); f("=ROWS(SEQUENCE(3,2))", 3);
{ const v = val(book({ Z1: "=RANDBETWEEN(1,6)" }), "Z1"); eq(v >= 1 && v <= 6 && Number.isInteger(v), true, "RANDBETWEEN range"); }
f("=SIN(0)", 0); f("=DEGREES(PI())", 180);

// ---- statistics
f("=AVERAGE(A1:A5)", 20); f("=AVERAGE(A4)", "#DIV/0!"); f('=AVERAGEIF(B1:B4,"apple",C1:C4)', 2); f("=COUNT(A1:A5)", 3); f("=COUNTA(A1:A5)", 4); f("=COUNTBLANK(A1:A5)", 1);
f('=COUNTIF(B1:B4,"apple")', 2); f('=COUNTIF(A1:A5,">=20")', 2); f('=COUNTIF(B1:B4,"<>apple")', 2); f('=COUNTIF(B1:B4,"?????")', 2); f('=COUNTIFS(B1:B4,"apple",C1:C4,">1")', 1);
f("=COUNTUNIQUE(B1:B4)", 3); f("=MAX(A1:A5)", 30); f("=MIN(A1:A5)", 10); f("=MAX(A4)", 0); f('=MAXIFS(C1:C4,B1:B4,"apple")', 3); f('=MINIFS(C1:C4,B1:B4,"apple")', 1);
f("=MEDIAN(C1:C4)", 2.5); f("=MODE(1,2,2,3)", 2); f("=STDEV(C1:C5)", Math.sqrt(2.5)); f("=STDEVP(C1:C5)", Math.sqrt(2)); f("=VAR(C1:C5)", 2.5); f("=VAR.P(C1:C5)", 2);
f("=LARGE(A1:A3,1)", 30); f("=SMALL(A1:A3,2)", 20); f("=RANK(20,A1:A3)", 2); f("=RANK(20,A1:A3,1)", 2); f("=RANK(10,A1:A3)", 3);
f("=PERCENTILE(C1:C5,0.25)", 2); f("=QUARTILE(C1:C5,2)", 3); f("=CORREL(A1:A3,C1:C3)", 1); f("=SLOPE(A1:A3,C1:C3)", 10); f("=INTERCEPT(A1:A3,C1:C3)", 0); f("=FORECAST(4,A1:A3,C1:C3)", 40);
f("=GEOMEAN(1,4)", 2); f("=HARMEAN(1,4)", 1.6); f("=NORM.S.DIST(0,TRUE)", 0.5); f("=ROUND(NORM.S.INV(0.975),2)", 1.96);

// ---- logical
f("=IF(A1>5,\"big\",\"small\")", "big"); f("=IF(FALSE,1)", false); f("=IF(A4=0,\"-\",1/A4)", "-"); f("=IF(TRUE,,1)", 0);
f("=IFS(A1>50,\"a\",A1>5,\"b\")", "b"); f("=IFS(A1>50,\"a\")", "#N/A"); f("=IFERROR(1/0,\"none\")", "none"); f("=IFERROR(5,0)", 5); f("=IFNA(NA(),\"x\")", "x"); f("=IFNA(1/0,\"x\")", "#DIV/0!");
f("=AND(TRUE,A1>5)", true); f("=OR(FALSE,A1>50)", false); f("=XOR(TRUE,TRUE)", false); f("=NOT(TRUE)", false); f("=AND(B1:B5)", true);
f('=SWITCH(2,1,"one",2,"two","other")', "two"); f('=SWITCH(9,1,"one","other")', "other"); f('=CHOOSE(2,"a","b","c")', "b"); f("=LET(x,A1*2,y,x+1,x*y)", 420);
f("=SUM(IF(A1:A3>15,A1:A3,0))", 50);

// ---- lookup
f("=VLOOKUP(\"Bob\",D2:E4,2,FALSE)", 95); f("=VLOOKUP(\"Zed\",D2:E4,2,FALSE)", "#N/A"); f("=VLOOKUP(\"b*\",D2:E4,2,FALSE)", 95); f("=VLOOKUP(25,A1:C3,3)", 2); f("=VLOOKUP(5,A1:C3,3)", "#N/A");
f("=VLOOKUP(\"Bob\",D2:E4,3,FALSE)", "#REF!"); f("=HLOOKUP(\"Score\",D1:E4,3,FALSE)", 95); f("=MATCH(30,A1:A3,0)", 3); f("=MATCH(25,A1:A3)", 2); f('=MATCH("cherry",B1:B4,0)', 4); f("=MATCH(25,A1:A3,0)", "#N/A");
f("=INDEX(D2:E4,2,2)", 95); f("=INDEX(A1:A3,2)", 20); f("=SUM(INDEX(D2:E4,0,2))", 255); f('=INDEX(E2:E4,MATCH("Cid",D2:D4,0))', 72);
f('=XLOOKUP("Cid",D2:D4,E2:E4)', 72); f('=XLOOKUP("Zed",D2:D4,E2:E4,"missing")', "missing"); f("=XLOOKUP(25,A1:A3,C1:C3,,-1)", 2); f("=XLOOKUP(25,A1:A3,C1:C3,,1)", 3); f('=XLOOKUP("apple",B1:B4,C1:C4,,0,-1)', 3);
f("=XMATCH(30,A1:A3)", 3); f("=LOOKUP(25,A1:A3,C1:C3)", 2); f("=ROW(C7)", 7); f("=COLUMN(C7)", 3); f("=ROW()", 1); f("=ROWS(A1:B9)", 9); f("=COLUMNS(A1:D1)", 4);
f("=ADDRESS(3,2)", "$B$3"); f("=ADDRESS(3,2,4)", "B3"); f('=INDIRECT("A"&2)', 20); f('=SUM(INDIRECT("A1:A3"))', 60); f("=SUM(OFFSET(A1,1,0,2,1))", 50); f("=OFFSET(A1,-1,0)", "#REF!");
f("=SUM(A1:INDEX(A1:A3,2))", 30);

// ---- spilling (dynamic arrays)
{
  const b = book({ ...DATA, H1: "=SEQUENCE(3)", I1: "=H2*10", J1: '=FILTER(D2:D4,E2:E4>80)', K1: "=SORT(E2:E4,1,FALSE)", L1: "=UNIQUE(B1:B4)", M1: "=TRANSPOSE(C1:C3)", N5: "=A1:A3*2" });
  eq(val(b, "H3"), 3, "SEQUENCE spills down"); eq(val(b, "I1"), 20, "a formula can read a spilled cell");
  eq(val(b, "J1"), "Ann", "FILTER first"); eq(val(b, "J2"), "Bob", "FILTER second"); eq(val(b, "J3"), null, "FILTER stops");
  eq(val(b, "K1"), 95, "SORT descending"); eq(val(b, "K3"), 72, "SORT last");
  eq(val(b, "L3"), "cherry", "UNIQUE"); eq(val(b, "L4"), null, "UNIQUE length");
  eq(val(b, "N1"), 2, "TRANSPOSE across"); eq(val(b, "N7"), 60, "array maths spills");
  eq(b.spillAt(0, 0, 7).h, 3, "spill anchor size");
  const blocked = book({ H1: "=SEQUENCE(3)", H2: "x" }); eq(val(blocked, "H1"), "#SPILL!", "blocked spill");
  eq(val(blocked, "H3"), null, "nothing spilled when blocked");
  eq(val(book({ H1: "=FILTER(A1:A3,A1:A3>99)", A1: "1", A2: "2", A3: "3" }), "H1"), "#CALC!", "FILTER finds nothing");
  eq(val(book({ H1: '=FILTER(A1:A3,A1:A3>99,"none")', A1: "1", A2: "2", A3: "3" }), "H1"), "none", "FILTER's if_empty, like Excel");
  eq(val(book({ H1: "=FILTER(A1:A3,A1:A3>1,B1:B3=\"y\")", A1: "1", A2: "2", A3: "3", B1: "y", B2: "n", B3: "y" }), "H1"), 3, "FILTER with two conditions, like Google Sheets");
  eq(val(book({ ...DATA, H1: "=LEN(B1:B3)" }), "H2"), 6, "functions lift over ranges");
  eq(val(book({ ...DATA, H1: "=SUM(H2:H3)", H2: "=SEQUENCE(1)" }), "H1"), 1, "sum over a spill source");
  eq(val(book({ ...DATA, H1: '=VSTACK({1;2},{3})' }), "H3"), 3, "VSTACK"); eq(val(book({ H1: '=HSTACK({1,2},{3})' }), "J1"), 3, "HSTACK");
  eq(val(book({ H1: '={1,2;3,4}' }), "I2"), 4, "array literal spills");
  eq(val(book({ ...DATA, H1: '=SPLIT("a,b,c",",")' }), "J1"), "c", "SPLIT spills across");
  eq(val(book({ ...DATA, H1: '=TOCOL(C1:C3)' }), "H3"), 3, "TOCOL");
}

// ---- text
f('=CONCATENATE("a",1,TRUE)', "a1TRUE"); f("=CONCAT(C1:C3)", "123"); f('=TEXTJOIN(", ",TRUE,B1:B4,A4)', "apple, banana, apple, cherry"); f('=JOIN("-",C1:C3)', "1-2-3");
f('=LEFT("hello",2)', "he"); f('=LEFT("hello")', "h"); f('=RIGHT("hello",3)', "llo"); f('=MID("hello",2,3)', "ell"); f('=LEN("hello")', 5); f('=UPPER("ab")', "AB"); f('=LOWER("AB")', "ab");
f('=PROPER("hello wORLD")', "Hello World"); f('=TRIM("  a   b ")', "a b"); f('=SUBSTITUTE("a-b-c","-","+")', "a+b+c"); f('=SUBSTITUTE("a-b-c","-","+",2)', "a-b+c"); f('=REPLACE("abcdef",2,3,"X")', "aXef");
f('=FIND("l","hello")', 3); f('=FIND("L","hello")', "#VALUE!"); f('=SEARCH("L","hello")', 3); f('=SEARCH("l?o","hello")', 3); f('=REPT("ab",3)', "ababab"); f('=EXACT("a","A")', false);
f("=CHAR(65)", "A"); f('=CODE("A")', 65); f('=TEXT(1234.5,"$#,##0.00")', "$1,234.50"); f('=TEXT(F1,"mmm d")', "Mar 15"); f('=TEXT(0.5,"0%")', "50%"); f('=VALUE("1,234")', 1234); f('=VALUE("abc")', "#VALUE!");
f('=TEXTBEFORE("a@b.com","@")', "a"); f('=TEXTAFTER("a@b.com","@")', "b.com"); f('=REGEXMATCH("abc123","\\d+")', true); f('=REGEXEXTRACT("abc123","\\d+")', "123"); f('=REGEXREPLACE("a1b2","\\d","#")', "a#b#");
f('=FIXED(1234.567,1)', "1,234.6"); f('=DOLLAR(-5)', "($5.00)"); f("=T(A1)", ""); f("=N(TRUE)", 1); f('="Total: "&TEXT(SUM(A1:A3),"#,##0")', "Total: 60");

// ---- dates
f("=DATE(2024,3,15)", 45366); f("=DATE(2024,13,1)", F.dateSerial(2025, 1, 1)); f("=DATE(2024,3,0)", F.dateSerial(2024, 2, 29)); f("=YEAR(F1)", 2024); f("=MONTH(F1)", 3); f("=DAY(F1)", 15);
f("=WEEKDAY(F1)", 6); f("=WEEKDAY(F1,2)", 5); f("=EDATE(F2,1)", F.dateSerial(2024, 2, 29)); f("=EOMONTH(F1,0)", F.dateSerial(2024, 3, 31)); f("=EOMONTH(F1,-1)", F.dateSerial(2024, 2, 29));
f("=DAYS(F1,F2)", 44); f("=F1-F2", 44); f('=DATEDIF(F2,F1,"M")', 1); f('=DATEDIF("2000-05-20","2024-03-15","Y")', 23); f('=DATEDIF(F2,F1,"D")', 44); f('=DATEDIF(F1,F2,"D")', "#NUM!");
f("=NETWORKDAYS(DATE(2024,3,11),DATE(2024,3,22))", 10); f("=WORKDAY(DATE(2024,3,15),1)", F.dateSerial(2024, 3, 18)); f("=HOUR(\"14:30\")", 14); f("=TIME(12,30,0)", 0.5 + 1 / 48);
f('=DATEVALUE("2024-03-15")', 45366); f("=WEEKNUM(DATE(2024,1,7))", 2); f("=ISOWEEKNUM(DATE(2024,1,1))", 1); f("=TODAY()=INT(NOW())", true); f("=YEAR(TODAY())>=2024", true);

// ---- info
f("=ISBLANK(A4)", true); f("=ISNUMBER(A1)", true); f("=ISTEXT(B1)", true); f("=ISERROR(1/0)", true); f("=ISNA(NA())", true); f("=ISERR(NA())", false); f("=ISEVEN(4)", true); f("=ISODD(4)", false);
f("=ISFORMULA(Z1)", true); f("=FORMULATEXT(Z1)", "=FORMULATEXT(Z1)"); f("=TYPE(\"a\")", 2); f("=ERROR.TYPE(1/0)", 2); f("=SHEET()", 1); f("=SHEETS()", 1);

// ---- finance (checked against Excel)
f("=ROUND(PMT(0.05/12,360,200000),2)", -1073.64); f("=ROUND(FV(0.05/12,120,-100),2)", 15528.23); f("=ROUND(PV(0.08/12,60,-500),2)", 24659.22); f("=ROUND(NPER(0.01,-100,1000),4)", 10.5886);
f("=ROUND(RATE(60,-500,24659.22)*12,4)", 0.08); f("=ROUND(NPV(0.1,-1000,300,400,500),2)", -19.12); f("=ROUND(IRR({-1000,300,400,500}),4)", 0.089);
f("=ROUND(IPMT(0.05/12,1,360,200000),2)", -833.33); f("=ROUND(PPMT(0.05/12,1,360,200000),2)", -240.31); f("=SLN(1000,100,9)", 100); f("=ROUND(EFFECT(0.12,12),6)", 0.126825);

// ---- references across sheets, circular references
{
  const b = book({ A1: "5", B1: "=Data!A1*2", C1: "='My Data'!B2+1", D1: "=D1+1", E1: "=E2", E2: "=E1" }, [{ name: "Data", cells: { A1: { v: 21 } } }, { name: "My Data", cells: { B2: { v: 1 } } }]);
  eq(val(b, "B1"), 42, "another sheet"); eq(val(b, "C1"), 2, "quoted sheet name"); eq(val(b, "D1"), "#REF!", "circular"); eq(val(b, "E1"), "#REF!", "two-cell circle");
  const col = book({ A1: "1", A2: "2", A9: "3", B1: "=SUM(A:A)", C5: "=SUM(1:1)" }); eq(val(col, "B1"), 6, "whole column"); eq(val(col, "C5"), 7, "whole row");
}

// ---- rewriting formulas (copy, insert/delete rows and columns, renaming sheets)
eq(F.shiftFormula("=A1+$B$2+C$3+$D4", 1, 1), "=B2+$B$2+D$3+$D5", "copy shifts relative parts");
eq(F.shiftFormula("=SUM(A1:A3)", 2, 0), "=SUM(A3:A5)", "copy a range");
eq(F.shiftFormula("=A1", -1, 0), "=#REF!", "copy off the top");
eq(F.shiftFormula('=A1&"A1"', 1, 0), '=A2&"A1"', "text in quotes is left alone");
eq(F.shiftFormula("=Sheet2!A1+SUM(B:B)", 1, 1), "=Sheet2!B2+SUM(C:C)", "sheets and whole columns");
eq(F.structFormula("=SUM(A1:A5)+A7", "Sheet1", "Sheet1", "r", 2, 1), "=SUM(A1:A6)+A8", "insert a row inside a range");
eq(F.structFormula("=SUM(A1:A5)+A7", "Sheet1", "Sheet1", "r", 1, -2), "=SUM(A1:A3)+A5", "delete rows inside a range");
eq(F.structFormula("=A3", "Sheet1", "Sheet1", "r", 2, -1), "=#REF!", "delete the row a formula points at");
eq(F.structFormula("=B1+Data!B1", "Sheet1", "Data", "c", 0, 1), "=B1+Data!C1", "only the changed sheet moves");
eq(F.structFormula("=$C$1", "Sheet1", "Sheet1", "c", 1, 1), "=$D$1", "absolute references move with inserted columns");
eq(F.renameSheetIn("=Data!A1+'Data'!B2", "Data", "My Data"), "='My Data'!A1+'My Data'!B2", "rename a sheet");
eq(F.dropSheetIn("=Data!A1+A1", "Data"), "=#REF!+A1", "delete a sheet");
eq(F.refsIn("=A1+SUM(B2:C3)").length, 2, "references for highlighting");
eq(F.refsIn("=A1+SUM(B2:C3)")[1].start, 8, "reference position");

// ---- every catalogued function is callable and has a description
for (const fn of F.FUNCTIONS) { if (!F.FN[fn.name] || !fn.desc) { fails++; console.log(`FAIL catalogue: ${fn.name}`); } else passes++; }

console.log(`${passes} passed, ${fails} failed`);
process.exit(fails ? 1 : 0);
