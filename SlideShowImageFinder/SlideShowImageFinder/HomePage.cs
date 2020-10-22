using HtmlAgilityPack;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Net;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Forms;

//About this program can be found under "Program.cs" in the solution explorer

namespace SlideShowImageFinder
{
    public partial class textTextbox : Form
    {
        public textTextbox()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
        }

        //This button takes in the title and the text of a users slide show
        //This event handler then finds all of the bold letters in the text field and puts them into a string
        //The event handler then converts the title and text fields into a usable url extension that will add onto http://www.searchscene.com/search?
        //Finally the event handler will then give this extension to the image generation page and then display the new page will images found using the search criteria
        private void generateButton_Click(object sender, EventArgs e)
        {
            string titleInput = titleTextBox.Text.ToString(); //Set the title to a string
            string textInput = richTextBoxText.Text.ToString();
            string boldTextInput;
            string boldFinder = "a";
            int counter = 0;

            //Used the code below to find which characters in the richTextBox were bold characters
            int end = richTextBoxText.SelectionStart;
            int start = richTextBoxText.SelectionLength;
            for (int i = 1; i < end; i++) //looped through every character in the richTextBox
            {
                richTextBoxText.SelectionStart = start + i;
                richTextBoxText.SelectionLength = 1;
                if (!richTextBoxText.SelectionFont.Bold) //If the char is not bold then move on
                {
                    richTextBoxText.SelectionStart = 0;
                    richTextBoxText.SelectionLength = 0;
                    richTextBoxText.SelectionStart = start;
                    richTextBoxText.SelectionLength = end;
                    richTextBoxText.Focus();
                }
                else if (counter == 0) //If the char is bold and the boldFinder string has nothing in it then set boldFinder to the first bold character
                {
                    boldFinder = richTextBoxText.Text[i].ToString();
                    counter++;
                }
                else //if the char is bold and the boldFinder string isn't empty then add the next bold character to the string
                {
                    boldFinder += richTextBoxText.Text[i].ToString();
                    if (richTextBoxText.Text[i + 1].ToString() == " ") //If the next character is a space then add a space to the boldFinder string and skip the next sequence
                    {
                        i++;
                        boldFinder += " ";
                    }
                    else if(richTextBoxText.Text[i + 1].ToString() == "\n") //If the next character is a newline then add a space to the boldFinder string and skip the next sequence
                    {
                        i++;
                        boldFinder += " ";
                    }
                }
            }
            
            boldTextInput = boldFinder;
            string fullInput = titleInput + " " + textInput; //Sets the full input given from the user into a string

            //parce through the full input and find space. if there is a space the replace with %20
            //%20 is used becuase the website url reads spaces as %20 for the image search
            string search = Regex.Replace(fullInput, @"\s+", "%20");
            string fullUrlExtension = search; //sets the full url extension to a string

            ImageGenerationPage igp = new ImageGenerationPage(fullUrlExtension, titleInput, textInput); //intializes the image generation page by providing the proper url extension, title, and text
            igp.Show(); //shows the image generation page

        }

        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {
            richTextBoxText.SwtichToBoldRegular();
        }
    }
    static class Helper
    {
        public static void SwtichToBoldRegular(this RichTextBox c)
        {
            if (c.SelectionFont.Style != FontStyle.Bold)
                c.SelectionFont = new Font(c.Font.FontFamily, c.Font.Size, FontStyle.Bold);
            else
                c.SelectionFont = new Font(c.Font.FontFamily, c.Font.Size, FontStyle.Regular);
        }
    }
}
