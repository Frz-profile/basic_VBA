# 🔢 Excel VBA Basics – 5 Simple Macros
This repository contains **5 beginner-friendly Excel VBA macro examples** to demonstrate basic operations such as math, conditional logic, string manipulation, and date functions. Each macro outputs its result to a specific cell in an Excel worksheet.

---

## 📄 Macro List
### 1. ➕ Add Two Numbers
```vba
Sub AddNumbers()
    Dim a As Double, b As Double
    a = 5
    b = 10
    Range("B2").Value = a + b
End Sub
```
📍 **Output:** Cell `B2`  
📝 **Description:** Adds two numbers and displays the sum.

---

### 2. 🔍 Check if a Number is Even
```vba
Sub CheckEven()
    Dim num As Integer
    num = 4
    Range("B5").Value = (num Mod 2 = 0)
End Sub
```
📍 **Output:** Cell `B5`  
📝 **Description:** Checks whether the number is even and displays `TRUE` or `FALSE`.

---

### 3. 📅 Show Today’s Date
```vba
Sub ShowToday()
    Range("B8").Value = Date
End Sub
```
📍 **Output:** Cell `B8`  
📝 **Description:** Displays the current date.

---

### 4. 🔠 Convert Text to Uppercase
```vba
Sub ConvertToUpper()
    Dim txt As String
    txt = "hello world"
    Range("B11").Value = UCase(txt)
End Sub
```
📍 **Output:** Cell `B11`  
📝 **Description:** Converts the given text to uppercase.

---

### 5. 🔢 Get Length of a String
```vba
Sub GetTextLength()
    Dim txt As String
    txt = "Excel VBA"
    Range("B14").Value = Len(txt)
End Sub
```
📍 **Output:** Cell `B14`  
📝 **Description:** Calculates and displays the number of characters in the string.

---

## 🧰 Requirements
- Microsoft Excel with Developer tab enabled
- Basic understanding of the VBA Editor (ALT + F11)

---

## 📦 How to Use
1. Open your Excel workbook.
2. Press `ALT + F11` to open the **VBA Editor**.
3. Insert a new **Module**.
4. Copy and paste any of the macros you want to try.
5. Run the macro by pressing `F5` or assign it to a button.
