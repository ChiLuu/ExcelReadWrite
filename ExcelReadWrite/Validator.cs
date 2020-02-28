using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;

namespace ExcelReadWrite
{
    // Repository of validation methods
    public static class Validator
    {
        // Check if Textbox data is present or else display corresponding message to user
        public static bool IsPresent(TextBox tb, string name)
        {
            bool valid = true;  // Return value. True till proven false.

            // If textbox's data is empty, set return value to false and display error message to user
            if (tb.Text == "")
            {
                valid = false;
                MessageBox.Show(name + " is required", "Input Error");
                tb.Focus();
            }

            return valid;
        }

        // Check if Textbox data is an Integer or else display corresponding message to user
        public static bool IsInt32(TextBox tb, string name)
        {
            bool valid = true;  // Return value. True till proven false.
            int value;          // Integer returned on parse method

            if (!Int32.TryParse(tb.Text, out value)) // Not an int, set return value to false
            {
                valid = false;
                MessageBox.Show(name + " must be a whole number.", "Input Error");
                tb.SelectAll();
                tb.Focus();
            }

            return valid;
        }

        // Check if Textbox data is a non-negative Integer or else display corresponding message to user
        public static bool IsNonNegativeInt32(TextBox tb, string name)
        {
            bool valid = true;  // Return value. True till proven false.
            int value;          // Integer returned on parse method

            if (!Int32.TryParse(tb.Text, out value)) // Not an int
            {
                valid = false;
                MessageBox.Show(name + " must be a whole number.", "Input Error");
                tb.SelectAll();
                tb.Focus();
            }
            else if (value < 0) // Negative, set return value to false
            {
                valid = false;
                MessageBox.Show(name + " must be positive or zero");
                tb.SelectAll();
                tb.Focus();
            }

            return valid;
        }

        // Check if Textbox data is a Double or else display corresponding message to user
        public static bool IsDouble(TextBox tb, string name)
        {
            bool valid = true;  // Return value. True till proven false.
            double value;       // Double returned on parse method

            if (!Double.TryParse(tb.Text, out value)) // Not an Double, set return value to false
            {
                valid = false;
                MessageBox.Show(name + " must be a number.", "Input Error");
                tb.SelectAll();
                tb.Focus();
            }

            return valid;
        }

        // Check if Textbox data is a non-negative Double or else display corresponding message to user
        public static bool IsNonNegativeDouble(TextBox tb, string name)
        {
            bool valid = true;  // Return value. True till proven false.
            double value;       // Double returned on parse method

            if (!Double.TryParse(tb.Text, out value)) // Not an Double, set return value to false
            {
                valid = false;
                MessageBox.Show(name + " must be a whole number.", "Input Error");
                tb.SelectAll();
                tb.Focus();
            }
            else if (value < 0) // Negative, set return value to false
            {
                valid = false;
                MessageBox.Show(name + " must be positive or zero");
                tb.SelectAll();
                tb.Focus();
            }

            return valid;
        }

        // Check if Textbox data is a Decimal or else display corresponding message to user
        public static bool IsDecimal(TextBox tb, string name)
        {
            bool valid = true;  // Return value. True till proven false.
            decimal value;      // Decimal returned on parse method

            if (!Decimal.TryParse(tb.Text, out value)) // Not an Decimal, set return value to false
            {
                valid = false;
                MessageBox.Show(name + " must be a number.", "Input Error");
                tb.SelectAll();
                tb.Focus();
            }

            return valid;
        }

        // Check if Textbox data is a non-negative Decimal or else display corresponding message to user
        public static bool IsNonNegativeDecimal(TextBox tb, string name)
        {
            bool valid = true;  // Return value. True till proven false.
            decimal value;      // Decimal returned on parse method

            if (!Decimal.TryParse(tb.Text, out value)) // Not an Decimal, set return value to false
            {
                valid = false;
                MessageBox.Show(name + " must be a whole number.", "Input Error");
                tb.SelectAll();
                tb.Focus();
            }
            else if (value < 0) // Negative, set return value to false
            {
                valid = false;
                MessageBox.Show(name + " must be positive or zero");
                tb.SelectAll();
                tb.Focus();
            }

            return valid;
        }

    }
}
