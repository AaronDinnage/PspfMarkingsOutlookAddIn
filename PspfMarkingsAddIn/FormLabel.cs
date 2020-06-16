using System;
using System.Windows.Forms;

namespace PspfMarkings
{
    public partial class FormLabel : Form
    {
        public string Selected
        {
            get
            {
                return comboBoxLabel.SelectedItem as string;
            } 
        }

        public FormLabel()
        {
            InitializeComponent();

            labelMessage.Text = Config.Current.RequireMeetingLabelMessage;
        }

        private void FormLabel_Load(object sender, EventArgs e)
        {
            comboBoxLabel.Items.Clear();
            foreach (var marking in Config.Current.ProtectiveMarkings)
                comboBoxLabel.Items.Add(marking.DisplayName);
        }

        private void comboBoxLabel_SelectedIndexChanged(object sender, EventArgs e)
        {
            buttonOK.Enabled = true;
        }
    }
}
