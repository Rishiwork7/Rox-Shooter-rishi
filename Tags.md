# 🏷️ Rox-Shooter Dynamic Tags Guide

In tags ka use aap apne Email Subject, Body, HTML template aur Filename ko dynamic banane ke liye kar sakte hain.

## 🚀 Available Tags

| Tag Syntax | Example | Description |
| :--- | :--- | :--- |
| **`$random[N]`** | `$random10` | Generates a mix of letters and numbers of length N. |
| **`$word[N]`** | `$word6` | Generates only random uppercase letters of length N. |
| **`$invoice_no[N]`** | `$invoice_no8` | Generates random digits ending with 1-2 letters. |
| **`$rand[N]`** | `$rand5` | Generates only random numbers of length N. |
| **`$number[N]`** | `$number7` | Same as `$rand`, generates N random digits. |
| **`$mail`** | `$mail` | Injects the recipient's email address. |
| **`$date`** | `$date` | Injects current date in DD-MM-YYYY format. |
| **`$tfn`** | `$tfn` | Injects the Toll-Free Number from settings. |

---

## 📝 Usage Examples

### 1. Subject Line
`Invoice Update for $mail - #$invoice_no10`
*Result: Invoice Update for user@gmail.com - #482910AK*

### 2. Email Body
`Hello, your order placed on $date is being processed. Ref: $random12`
*Result: Hello, your order placed on 06-05-2026 is being processed. Ref: 9J2L0P8Z1X4Q*

### 3. Filename (Custom Mode)
`Statement_$mail_$date`
*Result: Statement_user@gmail.com_06-05-2026.pdf*

---
**Note:** `[N]` is optional. If not provided, it defaults to a length of 6 characters.
