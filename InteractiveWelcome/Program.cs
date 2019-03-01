using System;
using System.Drawing;
using System.Windows.Forms;

// custom delegate
public delegate void Startdelegate();

class Eventdemo : Form
{
    // custom event
    public event Startdelegate StartEvent;

    public Eventdemo()
    {
        Button clickMe = new Button();

        clickMe.Parent = this;
        clickMe.Text = "Click Me";
        clickMe.Location = new Point(
            (ClientSize.Width - clickMe.Width) / 2,
            (ClientSize.Height - clickMe.Height) / 2);

        // an EventHandler delegate is assigned
        // to the button's Click event
        clickMe.Click += new EventHandler(OnClickMeClicked);

        // our custom "Startdelegate" delegate is assigned
        // to our custom "StartEvent" event.
        StartEvent += new Startdelegate(OnStartEvent);

        // fire our custom event
        StartEvent();
    }

    // this method is called when the "clickMe" button is pressed
    public void OnClickMeClicked(object sender, EventArgs ea)
    {
        MessageBox.Show("You Clicked My Button!");
    }

    // this method is called when the "StartEvent" Event is fired
    public void OnStartEvent()
    {
        MessageBox.Show("I Just Started!");
    }

    static void Main(string[] args)
    {
        Application.Run(new Eventdemo());
    }
}
//bubba 2 feathers