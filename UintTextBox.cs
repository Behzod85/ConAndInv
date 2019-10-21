using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Controls;
using System.Windows.Input;

namespace ConAndInv
{
    class UintTextBox : TextBox
    {
        static UintTextBox()
        {
        }

        protected override void OnTextInput(TextCompositionEventArgs e)
        {
            e.Handled = !UintCharChecker(e.Text);
            base.OnTextInput(e);
        }

        protected override void OnPreviewKeyDown(KeyEventArgs e)
        {
            e.Handled = (e.Key == Key.Space);
            base.OnPreviewKeyDown(e);
        }

        private bool UintCharChecker(string str)
        {
            foreach (char c in str)
            {
                    if (Char.IsNumber(c))
                    return true;
            }
            return false;
        }
    }
}

