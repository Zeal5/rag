#### **Basic Formulas**

| Formula | Structure | Explanation |
|---------|------------------|------------------------------------------------------------------|
| AVERAGE | =AVERAGE(A2:A10) | Returns a mathematical average of a given cell range |
| COUNT | =COUNT(A2:A10) | Returns the count of the numbers in given cell range |
| MAX | =MAX(A2:A10) | Finds the largest value in a given cell range |
| MEDIAN | =MEDIAN(A2:A10) | Returns the median value, or middle value, in a given cell range |
| MIN | =MIN(A2:A10) | Finds the smallest value in a given cell range |
| SUM | =SUM(A2:A10) | Totals numbers in a given cell range |

*Cell range A2:A10 is used above to indicate that each formula uses a cell range as it arguments* 

#### **Time Formulas**

| Formula | Structure | Explanation |
|---------|--------------------------------|---------------------------------------------------------------------|
| TODAY | =TODAY() | Volatile –takes no arguments –returns today's date |
| NOW | =NOW() | Volatile –takes no arguments –returns today's date and time |
| DATEDIF | =DATEDIF(Start Date, End Date, | Returns the number of years, months or days between two dates |
| | Unit) | Start Date –date furthest in the past |
| | | Unit could be "Y" for years, "M" for months or "D" for days |
| | | Units must be in double quotes |
| | | This formula is NOT in the function library |
| YEAR | =YEAR(Date) | Returns the year portion of date |
| | | Example =YEAR(7/16/2005) would return 2005 |
| MONTH | =MONTH(Date) | Returns the month portion of date |
| | | Example =MONTH(7/16/2005) would return 7 |
| DAY | =DAY(Date) | Returns the day portion of date |
| | | Example =DAY(7/16/2005) would return 16 |

*Use a time formula and get an answer you didn't expect? If you got a date and were expecting a number, remember to change the formatting from date to number. If you got a number and were expecting a date, change the formatting to date.* 

#### **Logical Formulas**

| Formula | Structure | Explanation |
| IF | =IF(Logical Test, TRUE, FALSE) | Evaluates the statement in the logical test to determine if it isTRUE or FALSEA Logical testcomparesthe value of one cell to another, or a cell value to aconstant value, using a comparison operator such as:oEqual =Less than <ooGreater than >oLess than or equal to <=oGreater than or equal to >=oNot equal to <>TRUE–this part of the IF function will only execute when the logical test is TRUEFALSE–this part of the IF function will only execute when the logical test is FALSEBoth TRUE and FALSE can be a word, a formula or a constant valueTo return a blank cell use two double quotes "" |
| OR | =OR(Logical Test 1, Logical Test2, ) | This formula can only return TRUE or FALSEOnly one test in the group must return TRUE for the formula to return TRUE |
| AND | =AND(Logical Test 1, LogicalTest 2, ) | This formula can only return TRUE or FALSEOnly one test in the group must return FALSE for the formula to return FALSE |

### **Lookup Formulas**

| Formula | Structure | Explanation |
|---------|---------------------------------|----------------------------------------------------------------------------------------------------------------------|
| VLOOKUP | =VLOOKUP (Lookup Value, Table | Lookup Value-What the function is looking for in the table array |
| | Array, Col Index, Range Lookup) | Table Array-The table defined as a cell range |
| | | Col Index-The column in the table that forms the return |
| | | Range lookup-False for exact match, True or blank for near match |
| | | Notes: |
| | | oA vlookup can only search vertically through the left most column of atable array for near or exact matches |
| | | In most cases you will want to use absolute cell referencing whenoindicating a table array |
| | | If you omit the Range Lookup, Excel will assume "True" and look for aonear match |

#### **Financial Formulas**

| Formula | Structure | Explanation |
|---------|-------------------------------------|------------------------------------------------------------------------------------------|
| PMT | = PMT(rate, nper, pv, [fv], [type]) | rate –Annual Percentage Rate |
| | | Divide yearly rate by 12 monthso |
| | | oREQUIRED |
| | | nper –number of periods |
| | | oHowmany months are in the loan? (12 months/year)* length of loan in |
| | | years= length of loan in months |
| | | REQUIREDo |
| | | pv –present value |
| | | oamount of loan –this is the amountthat wasborrowed |
| | | REQUIREDo |
| | | [fv]and [type]are in square brackets because they're optionalarguments. |
| | | ofv –future value -\$0 if the loan is paid if full |
| | | Excel assumes zero if omitted |
| | | otype – |
| | | 0 (zero) means the payments are due at the end of each period, |
| | | 1 means they're due at the beginning |
| | | Excel assumes your payments are due at the end of the period if |
| | | omitted |
| | | |

### **Statistical Formulas**

| Formula | Structure | Explanation |
| SUMIF | = SUMIF(range, criteria,[sum_range]) | SUMIF will return at total based on one criteriaRange -Required. The range of cells that you want evaluated by criteria.Criteria -Required.The criteria in the form of a number, expression, a cell reference, text, or afunction that defines which cells will be added.For example, criteria can be expressed as 32, ">32", B5, "32", "apples", orTODAY().Important: Any text criteria or any criteria that includes logical or mathematicalsymbols must be enclosed in double quotation marks ("). If the criteria isnumeric, double quotation marks are not required.Sum_Range Optional.The actual cells to add, if you want to add cells other than thoseospecified in the range argument.oIf the sum_range argument is omitted, Excel adds the cells that arespecified in the range argument (the same cells to which the criteria isapplied). |
| SUMIFS | =SUMIFS(sum_range,criteria_range1, criteria1,[criteria_range2, criteria2], ) | SUMIFS will return at total based on more than one criteriaSum_Range -The range of cells to sumCriteria_range1 –RequiredoThe rangethat is tested using Criteria1Criteria1 –RequiredThe first value tested for –this must be a match to be included in theototalCriteria_range2 –OptionaloThe range that is tested using Criteria2Criteria2 –OptionalThe second value tested for –this must be a match to be included in theototalAnd so on… |

| Formula | Structure | Explanation |
| COUNTIF | =COUNTIF(range, criteria) | Answers the question "How many of something(criteria) exist within specific setof cells (range)?=COUNTIF(Where do you want to look?, What do you want to look for?)Range –Required –Sets cells to be included in the countCriteria –Required –Tells formula what to look for |
| COUNTIFS | =COUNTIFS(criteria_range1,criteria1, [criteria_range2,criteria2]…) | criteria_range1 –Required.The first range in which to evaluate the associated criteria.ocriteria1 –Required.oThe criteria in the form of a number,expression, cell reference, or textthat define which cells will be counted. For example, criteria can beexpressed as:32,">32",B4,"apples", or"32"criteria_range2, criteria2,oOptional.oAdditional ranges and their associated criteria. Upto 127 range/criteriapairs are allowed.Important:oEach additional range must have the same number of rows andcolumns as the criteria_range1 argument.The ranges do not have to be adjacent to each othero |

| | =AVERAGEIF(range, criteria, | |
| AVERAGEIF | [average_range]) | Returns the average (arithmetic mean) of all the cells in a range that meet agiven criteriaRange –Required.One or more cells to average, including numbers or names, arrays, ororeferences that contain numbers.Criteria–Required.oThe criteria in the form of a number, expression, cell reference, or textthat defines which cells are averaged. For example, criteria can beexpressed as:32,">32",B4,"apples", or"32"Average_range –Optional.oThe actual set of cells to average.If omitted, range is usedo |

| Formula | Structure | Explanation |
| AVERAGEIFS | =AVERAGEIFS(average_range,criteria_range1, criteria1,[criteria_range2, criteria2], ) | Returns the average (arithmetic mean) of all cells that meet multiple criteriaAverage_range –RequiredOne or more cells to average, including numbers or names, arrays, ororeferences that contain numbers.Criteria_range1–Required, subsequent criteria_ranges are optional (up to 127ranges)criteria_range2, criteria_range3,…–OptionalCriteria1–Required,subsequent criteriaare optionaloCriteria can be in the form of a number, expression, cell reference, or textthat define which cells will be averaged. For example, criteria can beexpressed as:32,">32",B4,"apples", or"32"criteria2, criteria3,If you have a Criteria_Range, you must have a corresponding Criteria |
