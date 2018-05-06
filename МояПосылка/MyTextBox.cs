using System.Drawing;

namespace МояПосылка
{
    class MyTextBox : System.Windows.Forms.TextBox
    {
        public MyTextBox()
        {
            SetStyle(System.Windows.Forms.ControlStyles.UserPaint,
                     true);
        }

        protected override void
           OnPaint(System.Windows.Forms.PaintEventArgs e)
        {
            base.OnPaint(e);

            Rectangle rect = new Rectangle(e.ClipRectangle.X,
                e.ClipRectangle.Y,
                e.ClipRectangle.Width - 1,
                e.ClipRectangle.Height - 1);

            e.Graphics.DrawRectangle(new Pen(Color.DarkGreen, 1),
                rect);
        }
    }
}
