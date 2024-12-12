from pyrevit import forms
import os
from System.Windows import Window, WindowStartupLocation
from System.Windows.Controls import StackPanel, TextBlock, Button, Image
from System.Windows.Media.Imaging import BitmapImage
from System.Windows import HorizontalAlignment, SizeToContent, ResizeMode, TextWrapping
from System import Uri, UriKind

ITALIJAN_LENGTH_CM = 7.5
AVG_ITALIAN_HEIGHT = {
    'Musko': 175.0,
    'Zensko': 162.0
}

# Ask gender
gender = forms.SelectFromList.show(
    ['Musko', 'Zensko'],
    title='Izaberite Pol',
    width=300,
    height=200,
    button_name='OK'
)
if not gender:
    raise SystemExit

# Ask height
height_str = forms.ask_for_string(
    prompt="Tvoja visina u CM:",
    title="Tvoja visina",
    default=""
)
if not height_str:
    raise SystemExit

try:
    user_height = float(height_str)
except:
    forms.alert("Uneta visina nije validna.")
    raise SystemExit

how_many_italijan = user_height / ITALIJAN_LENGTH_CM
spacing_cm = 80.0
how_many_for_fix = round(user_height / spacing_cm)
how_many_real_italians = user_height / AVG_ITALIAN_HEIGHT[gender]

script_dir = os.path.dirname(__file__)
image_path = os.path.join(script_dir, "OK.png")

if gender == 'Musko':
    height_phrase = "visok"
else:
    height_phrase = "visoka"

msg = ("Tvoja visina: {} cm\n\n"
       "Ti si {:.2f} Italijana {}!!!\n\n"
       "Da bi pricvrstili holker za tebe potrebno je {} Italijana\n\n"
       "{} si {:.2f} Italijana (pravih onih iz Italije)\n").format(
           user_height,
           how_many_italijan,
           height_phrase,
           how_many_for_fix,
           height_phrase.capitalize(),
           how_many_real_italians
       )

# Simplify image display logic
if not os.path.exists(image_path):
    forms.alert("Slika nije pronadjena: {}".format(image_path))
    raise SystemExit

def show_results_with_image(message, img_path):
    wnd = Window()
    wnd.Title = "Rezultat"
    wnd.WindowStartupLocation = WindowStartupLocation.CenterScreen
    wnd.Width = 500  # Set custom width
    wnd.Height = 500  # Set custom height
    wnd.ResizeMode = ResizeMode.NoResize

    stack_panel = StackPanel()

    message_text = TextBlock()
    message_text.Text = message
    message_text.TextWrapping = TextWrapping.Wrap
    stack_panel.Children.Add(message_text)

    if os.path.exists(img_path):
        image = Image()
        bmp = BitmapImage()
        bmp.BeginInit()
        bmp.UriSource = Uri(img_path, UriKind.Absolute)
        bmp.EndInit()
        image.Source = bmp
        image.Height = 300  # Adjust image height if needed
        stack_panel.Children.Add(image)

    ok_button = Button()
    ok_button.Content = "OK"
    ok_button.HorizontalAlignment = HorizontalAlignment.Center
    ok_button.Click += lambda s, e: wnd.Close()
    stack_panel.Children.Add(ok_button)

    wnd.Content = stack_panel
    wnd.ShowDialog()

show_results_with_image(msg, image_path)
