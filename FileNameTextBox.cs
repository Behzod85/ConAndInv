using System;
using System.IO;
using System.Linq;
using System.Windows.Controls;
using System.Windows.Input;

namespace ConAndInv
{
    class FileNameTextBox : TextBox
    {
        static FileNameTextBox()
        {
        }

        protected override void OnTextInput(TextCompositionEventArgs e)
        {
            e.Handled = !FileNameValidCharNotEntered(e.Text.ToCharArray()[0]);
            base.OnTextInput(e);
        }

        protected override void OnPreviewKeyDown(KeyEventArgs e)
        {
            e.Handled = (e.Key == Key.Space);
            base.OnPreviewKeyDown(e);
        }

        private bool FileNameValidCharNotEntered(char c)
        {
            if (Path.GetInvalidFileNameChars().Contains(c) ||
            Path.GetInvalidPathChars().Contains(c))
            {
                return false;
            }
            return true;
        }
    }
}
