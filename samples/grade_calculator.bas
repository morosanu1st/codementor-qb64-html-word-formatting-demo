' ==========================================================
' Grade Calculator - a step-by-step QB64 demo program
' Each "STEP n:" comment marks one box in the flow diagram.
' ==========================================================

' STEP 1: Declare variables
DIM scores(1 TO 5) AS INTEGER
DIM i AS INTEGER
DIM total AS INTEGER
DIM average AS SINGLE
DIM grade AS STRING

' STEP 2: Read the five test scores
DATA 78,92,85,64,88
FOR i = 1 TO 5
    READ scores(i)
NEXT i

' STEP 3: Compute the total and the average
total = 0
FOR i = 1 TO 5
    total = total + scores(i)
NEXT i
average = total / 5

' STEP 4: Convert the average to a letter grade
SELECT CASE average
    CASE IS >= 90: grade = "A"
    CASE IS >= 80: grade = "B"
    CASE IS >= 70: grade = "C"
    CASE IS >= 60: grade = "D"
    CASE ELSE: grade = "F"
END SELECT

' STEP 5: Print the report
CLS
PRINT "Student Score Report"
PRINT "--------------------"
FOR i = 1 TO 5
    PRINT USING "Test ##: ###"; i; scores(i)
NEXT i
PRINT "--------------------"
PRINT USING "Total:   ####"; total
PRINT USING "Average: ###.##"; average
PRINT "Letter grade: "; grade
END
