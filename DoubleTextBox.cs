using System;
using System.Windows.Controls;
using System.Windows.Input;

namespace ConAndInv
{
    class DoubleTextBox : TextBox
    {
        static DoubleTextBox()
        {
        }

        protected override void OnTextInput(TextCompositionEventArgs e)
        {
            var s = (DoubleTextBox)e.Source;
            e.Handled = !DoubleCharChecker(e.Text, s.Text);
            //e.Handled = (e.Text.Equals('.') && (s.Text.IndexOf('.') != -1));
            base.OnTextInput(e);
        }

        protected override void OnPreviewKeyDown(KeyEventArgs e)
        {

            e.Handled = (e.Key == Key.Space);
            base.OnPreviewKeyDown(e);
        }

        private bool DoubleCharChecker(string str, string text)
        {
            foreach (char c in str)
            {
                if (c.Equals('-'))
                    return text == "";

                else if (c.Equals('.'))
                {

                    return text.IndexOf(',') == -1 && text.IndexOf('.') == -1 && text != "";
                }
                else if (c.Equals(','))
                {

                    return text.IndexOf('.') == -1 && text.IndexOf(',') == -1 && text != "";
                }


                else if (Char.IsNumber(c))
                    return true;
            }
            return false;
        }
    }
}
