using HtmlAgilityPack;
using Syncfusion.Presentation;
using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
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
    public partial class ImageGenerationPage : Form
    {
        private string pptTitle { get; set; }
        private string pptText { get; set; }
        private string searchUrl { get; set; }
        private int indexLimit { get; set; }
        private int index { get; set; }
        private string requestUrl { get; set; }
        private int img1SaveIndex {get;set;}
        private int img2SaveIndex { get; set; }
        private int img3SaveIndex { get; set; }


        public ImageGenerationPage(string search, string title, string text)
        {
            pptTitle = title;
            pptText = text;
            searchUrl = search;

            //I am only allowing for 21 images to show up in the search
            //If I do more than this then the program will become slower
            //I did this because the user wants to get the images as fast as possible
            indexLimit = 21;
            img1SaveIndex = 0;
            img2SaveIndex = 0;
            img3SaveIndex = 0;
            index = 0;

            //The Url the program is going to go to and search
            //using searchscene.com because it was the first search engine I could find that let me take image urls from it
            requestUrl =
              string.Format("http://www.searchscene.com/search?" +
               "q={0}&searchType=images", searchUrl);

            InitializeComponent();
        }

        //Generates the first 3 images found on the web page and puts them into picture boxes
        private void ImageGenerationPage_Load(object sender, EventArgs e)
        {
            //On load the application loads the html and then loads all url links that go to images
            HtmlAgilityPack.HtmlDocument document = new HtmlAgilityPack.HtmlDocument();

            document = new HtmlWeb().Load(requestUrl);
            //LINQ query that makes a list of links forr image urls
            var links = document.DocumentNode.Descendants("img")
                                .Select(a => a.GetAttributeValue("src", null))
                                .Where(s => !String.IsNullOrEmpty(s));

            //This allowed me to call from the list
            var urls = links.ToList();
            //Every Even element in the list is not an image so I removed the first element of the list
            urls.RemoveAt(0);
            var counter = 0;
            string imageUrl;
            byte[] imageBytes;
            HttpWebRequest imageRequest;
            WebResponse imageResponse;
            
            //Here I am using the web client to get the image url and the put the url into a byte array
            //where I then put the byte array into a stream which lets me set the image for each picture box
            using (WebClient webClient = new WebClient())
            {
                //I only need 3 images to display on the page at once so I get the first
                //3 images in the url list to display (Note I have to skip every second url)
                for (int i = 0; i < index+6; i+=2)
                {
                    imageUrl = urls[i];

                    imageRequest = (HttpWebRequest)WebRequest.Create(imageUrl);
                    imageResponse = imageRequest.GetResponse();

                    Stream responseStream = imageResponse.GetResponseStream();

                    using (BinaryReader br = new BinaryReader(responseStream))
                    {
                        imageBytes = br.ReadBytes(500000);
                        br.Close();
                    }
                    responseStream.Close();
                    imageResponse.Close();

                    using (MemoryStream stream = new MemoryStream(imageBytes))
                    {
                        //Set the image url to the picture boxes
                        if (counter == 0)
                        {
                            pictureBox1.Image = Image.FromStream(stream);
                            pictureBox1.SizeMode = PictureBoxSizeMode.StretchImage; //Setting the image fit to the size of the picture box
                            img1SaveIndex = index; //Keeping track of image index if I need to download it from the list later
                        }
                        else if (counter == 1)
                        {
                            pictureBox2.Image = Image.FromStream(stream);
                            pictureBox2.SizeMode = PictureBoxSizeMode.StretchImage;
                            img2SaveIndex = index+1;
                        }
                        else if (counter == 2)
                        {
                            pictureBox3.Image = Image.FromStream(stream);
                            pictureBox3.SizeMode = PictureBoxSizeMode.StretchImage;
                            img3SaveIndex = index+3;
                        }

                        if (counter < 3)
                        {
                            //counting index to keep track of which image I am at
                            index++;
                        }
                        //Using counter to know which picture box needs an image next
                        counter++;
                    }
                }
            }
            index++;
        }

        //Generates the next 3 images found on the web page and puts them into picture boxes
        private void nextButton_Click(object sender, EventArgs e)
        {
            //If the max index limit of 21 images (42 indexes) has been reached then tell the user
            if (index + 2 >= indexLimit * 2)
            {
                MessageBox.Show("No More Images!", "No More Images", MessageBoxButtons.OK);
            }
            else
            {
                //Get the html and image urls again and put them into a list
                HtmlAgilityPack.HtmlDocument document = new HtmlAgilityPack.HtmlDocument();

                document = new HtmlWeb().Load(requestUrl);

                var links = document.DocumentNode.Descendants("img")
                                    .Select(a => a.GetAttributeValue("src", null))
                                    .Where(s => !String.IsNullOrEmpty(s));

                var urls = links.ToList();

                urls.RemoveAt(0);
                var counter = 0;
                string imageUrl;
                byte[] imageBytes;
                HttpWebRequest imageRequest;
                WebResponse imageResponse;

                //Here I am using the web client to get the image url and the put the url into a byte array
                //where I then put the byte array into a stream which lets me set the image for each picture box
                using (WebClient webClient = new WebClient())
                {
                    //I only need 3 images to display on the page at once so I get the next
                    //3 images in the url list to display (Note I have to skip every second url)
                    for (int i = index + 2; i < index+6; i += 2)
                    {
                        imageUrl = urls[i];

                        imageRequest = (HttpWebRequest)WebRequest.Create(imageUrl);
                        imageResponse = imageRequest.GetResponse();

                        Stream responseStream = imageResponse.GetResponseStream();

                        using (BinaryReader br = new BinaryReader(responseStream))
                        {
                            imageBytes = br.ReadBytes(500000);
                            br.Close();
                        }
                        responseStream.Close();
                        imageResponse.Close();

                        using (MemoryStream stream = new MemoryStream(imageBytes))
                        {
                            //Set the image url to the picture boxes
                            if (counter == 0)
                            {
                                pictureBox1.Image = Image.FromStream(stream);
                                pictureBox1.SizeMode = PictureBoxSizeMode.StretchImage;
                                img1SaveIndex = index + 2; //Keeping track of image index if I need to download it from the list later
                            }
                            else if (counter == 1)
                            {
                                pictureBox2.Image = Image.FromStream(stream);
                                pictureBox2.SizeMode = PictureBoxSizeMode.StretchImage;
                                img2SaveIndex = index + 3;
                            }
                            else if (counter == 2)
                            {
                                pictureBox3.Image = Image.FromStream(stream);
                                pictureBox3.SizeMode = PictureBoxSizeMode.StretchImage;
                                img3SaveIndex = index + 4;
                            }

                            if (counter < 3)
                            {
                                //counting index to keep track of which image I am at
                                index++;
                            }
                            //Using counter to know which picture box needs an image next
                            counter++;
                        }
                    }
                }
                index += 3;
            }
        }

        //By double clicking the 3rd picture box you will be prompted to save the image
        //If yes then a file explorer will pop up asking where to store the image
        //The image will then be saved as a jpeg when the file location has been chosen
        //If you do not want to save then the message box will close
        private void pictureBox3_DoubleClick(object sender, EventArgs e)
        {
            //Message box initialization
            string message = "Do you want to save the selected image?";
            string title = "Save Image";
            MessageBoxButtons buttons = MessageBoxButtons.YesNo;
            DialogResult result = MessageBox.Show(message, title, buttons);

            //Save File Dialog being initiliazed
            SaveFileDialog save = new SaveFileDialog();
            save.Filter = "JPeg Image|*.jpg";
            save.Title = "Save an Image File";

            //getting the list of image urls again and then indexing to where the current 3rd picture box image is
            HtmlAgilityPack.HtmlDocument document = new HtmlAgilityPack.HtmlDocument();

            document = new HtmlWeb().Load(requestUrl);

            var links = document.DocumentNode.Descendants("img")
                                .Select(a => a.GetAttributeValue("src", null))
                                .Where(s => !String.IsNullOrEmpty(s));

            var urls = links.ToList();

            urls.RemoveAt(0);

            if (result == DialogResult.Yes)
            {
                save.ShowDialog();

                if (save.FileName != "")
                {
                    // Saves the Image via a FileStream created by the OpenFile method.
                    System.IO.FileStream fs =
                        (System.IO.FileStream)save.OpenFile();
                    // Saves the Image in the appropriate ImageFormat based upon the
                    // File type selected in the dialog box.
                    // NOTE that the FilterIndex property is one-based.
                    switch (save.FilterIndex)
                    {
                        case 1:
                            //The code below sets the image that is being indexed to a Bitmap which is then saved to the file system
                            using (WebClient webClient = new WebClient())
                            {
                                string imageUrl = urls[img3SaveIndex];

                                byte[] imageBytes;
                                HttpWebRequest imageRequest = (HttpWebRequest)WebRequest.Create(imageUrl);
                                WebResponse imageResponse = imageRequest.GetResponse();

                                Stream responseStream = imageResponse.GetResponseStream();

                                using (BinaryReader br = new BinaryReader(responseStream))
                                {
                                    imageBytes = br.ReadBytes(500000);
                                    br.Close();
                                }
                                responseStream.Close();
                                imageResponse.Close();

                                using (MemoryStream stream = new MemoryStream(imageBytes))
                                {
                                    using (Bitmap bmb = (Bitmap)Image.FromStream(stream).Clone())
                                    {
                                        bmb.Save(fs, System.Drawing.Imaging.ImageFormat.Jpeg);
                                    }
                                }
                            }
                            break;
                    }

                    fs.Close();
                }
            }
            else
            {
                //Close Box  
            }
        }

        //By double clicking the 2nd picture box you will be prompted to save the image
        //If yes then a file explorer will pop up asking where to store the image
        //The image will then be saved as a jpeg when the file location has been chosen
        //If you do not want to save then the message box will close
        private void pictureBox2_Click(object sender, EventArgs e)
        {
            string message = "Do you want to save the selected image?";
            string title = "Save Image";
            MessageBoxButtons buttons = MessageBoxButtons.YesNo;
            DialogResult result = MessageBox.Show(message, title, buttons);

            SaveFileDialog save = new SaveFileDialog();
            save.Filter = "JPeg Image|*.jpg";
            save.Title = "Save an Image File";

            HtmlAgilityPack.HtmlDocument document = new HtmlAgilityPack.HtmlDocument();

            document = new HtmlWeb().Load(requestUrl);

            var links = document.DocumentNode.Descendants("img")
                                .Select(a => a.GetAttributeValue("src", null))
                                .Where(s => !String.IsNullOrEmpty(s));

            var urls = links.ToList();

            urls.RemoveAt(0);

            if (result == DialogResult.Yes)
            {
                save.ShowDialog();

                if (save.FileName != "")
                {
                    // Saves the Image via a FileStream created by the OpenFile method.
                    System.IO.FileStream fs =
                        (System.IO.FileStream)save.OpenFile();
                    // Saves the Image in the appropriate ImageFormat based upon the
                    // File type selected in the dialog box.
                    // NOTE that the FilterIndex property is one-based.
                    switch (save.FilterIndex)
                    {
                        case 1:
                            using (WebClient webClient = new WebClient())
                            {
                                string imageUrl = urls[img2SaveIndex];

                                byte[] imageBytes;
                                HttpWebRequest imageRequest = (HttpWebRequest)WebRequest.Create(imageUrl);
                                WebResponse imageResponse = imageRequest.GetResponse();

                                Stream responseStream = imageResponse.GetResponseStream();

                                using (BinaryReader br = new BinaryReader(responseStream))
                                {
                                    imageBytes = br.ReadBytes(500000);
                                    br.Close();
                                }
                                responseStream.Close();
                                imageResponse.Close();

                                using (MemoryStream stream = new MemoryStream(imageBytes))
                                {
                                    using (Bitmap bmb = (Bitmap)Image.FromStream(stream).Clone())
                                    {
                                        bmb.Save(fs, System.Drawing.Imaging.ImageFormat.Jpeg);
                                    }
                                }
                            }
                            break;
                    }

                    fs.Close();
                }
            }
            else
            {
                //Close Box  
            }
        }

        //By double clicking the 1st picture box you will be prompted to save the image
        //If yes then a file explorer will pop up asking where to store the image
        //The image will then be saved as a jpeg when the file location has been chosen
        //If you do not want to save then the message box will close
        private void pictureBox1_Click(object sender, EventArgs e)
        {
            string message = "Do you want to save the selected image?";
            string title = "Save Image";
            MessageBoxButtons buttons = MessageBoxButtons.YesNo;
            DialogResult result = MessageBox.Show(message, title, buttons);

            SaveFileDialog save = new SaveFileDialog();
            save.Filter = "JPeg Image|*.jpg";
            save.Title = "Save an Image File";

            HtmlAgilityPack.HtmlDocument document = new HtmlAgilityPack.HtmlDocument();

            document = new HtmlWeb().Load(requestUrl);

            var links = document.DocumentNode.Descendants("img")
                                .Select(a => a.GetAttributeValue("src", null))
                                .Where(s => !String.IsNullOrEmpty(s));

            var urls = links.ToList();

            urls.RemoveAt(0);

            if (result == DialogResult.Yes)
            {
                save.ShowDialog();

                if (save.FileName != "")
                {
                    // Saves the Image via a FileStream created by the OpenFile method.
                    System.IO.FileStream fs =
                        (System.IO.FileStream)save.OpenFile();
                    // Saves the Image in the appropriate ImageFormat based upon the
                    // File type selected in the dialog box.
                    // NOTE that the FilterIndex property is one-based.
                    switch (save.FilterIndex)
                    {
                        case 1:
                            using (WebClient webClient = new WebClient())
                            {
                                string imageUrl = urls[img1SaveIndex];

                                byte[] imageBytes;
                                HttpWebRequest imageRequest = (HttpWebRequest)WebRequest.Create(imageUrl);
                                WebResponse imageResponse = imageRequest.GetResponse();

                                Stream responseStream = imageResponse.GetResponseStream();

                                using (BinaryReader br = new BinaryReader(responseStream))
                                {
                                    imageBytes = br.ReadBytes(500000);
                                    br.Close();
                                }
                                responseStream.Close();
                                imageResponse.Close();

                                using (MemoryStream stream = new MemoryStream(imageBytes))
                                {
                                    using (Bitmap bmb = (Bitmap)Image.FromStream(stream).Clone())
                                    {
                                        bmb.Save(fs, System.Drawing.Imaging.ImageFormat.Jpeg);
                                    }
                                }
                            }
                            break;
                    }

                    fs.Close();
                }
            }
            else
            {
                //Close Box  
            }
        }

        private void makePPTButton_Click(object sender, EventArgs e)
        {
            IPresentation pptxDoc = Presentation.Create(); //Creates new powerpoint

            ISlide slide = pptxDoc.Slides.Add(SlideLayoutType.TitleAndContent); //Creates new slide

            //Add title content to the slide by accessing the title placeholder of the TitleOnly layout-slide
            IShape titleShape = slide.Shapes[0] as IShape;
            titleShape.TextBody.AddParagraph(pptTitle).HorizontalAlignment = HorizontalAlignmentType.Center;

            //Adds content to the text box
            IShape descriptionShape = slide.Shapes[1] as IShape;
            descriptionShape.TextBody.AddParagraph(pptText);

            pptxDoc.Save("Sample.pptx");
        }
    }
}
