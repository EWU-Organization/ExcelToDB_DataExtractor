using System.Windows;
using System.Windows.Media.Animation;

namespace ExcelToDB_DataExtractor
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        private static MainWindow? _instance;
        public static MainWindow Instance
        {
            get
            {
                if (_instance == null)
                {
                    _instance = new MainWindow();
                }
                return _instance;
            }
        }
        public MainWindow()
        {
            InitializeComponent(); //340
            _instance = this;
            //WindowAnimation();
        }

        public void WindowAnimation(double from, double to)
        {
            // Create a new double animation
            DoubleAnimation animation = new DoubleAnimation();

            // Set the start value of the animation
            animation.From = from; // 285;

            // Set the end value of the animation
            animation.To = to; // 340;

            // Set the duration of the animation
            animation.Duration = new Duration(TimeSpan.FromSeconds(0.3));

            // Set the auto reverse property to true
            //animation.AutoReverse = true;

            // Set the repeat behavior to forever
            //animation.RepeatBehavior = RepeatBehavior.Forever;

            // Apply the animation to the height property of the window
            this.BeginAnimation(Window.HeightProperty, animation);
        }
    }
}