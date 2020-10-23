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
        private ArrayList imageToPptIndex { get; set; }
        private List<string> imageList { get; set; }


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
            imageToPptIndex = new ArrayList();

            //The Url the program is going to go to and search
            //using searchscene.com because it was the first search engine I could find that let me take image urls from it
            requestUrl =
              string.Format("http://www.searchscene.com/search?" +
               "q={0}&searchType=images", searchUrl);

            //On load the application loads the html and then loads all url links that go to images
            HtmlAgilityPack.HtmlDocument document = new HtmlAgilityPack.HtmlDocument();

            document = new HtmlWeb().Load(requestUrl);
            //LINQ query that makes a list of links forr image urls
            var links = document.DocumentNode.Descendants("img")
                                .Select(a => a.GetAttributeValue("src", null))
                                .Where(s => !String.IsNullOrEmpty(s));

            //This allowed me to call from the list
            imageList = links.ToList();
            imageList.RemoveAt(0);
            //Every Even element in the list is not an image so I removed the first element of the list
            for (int i = 0; i < imageList.Count; i++)
            {
                if (imageList[i].Contains("https://feed.cf"))
                {
                    imageList.RemoveAt(i);
                }
            }
            


            InitializeComponent();
        }

        //Generates the first 3 images found on the web page and puts them into picture boxes
        private void ImageGenerationPage_Load(object sender, EventArgs e)
        {
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
                for (int i = 0; i < 3; i++)
                {
                    imageUrl = imageList[i];

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
                            img2SaveIndex = index;
                        }
                        else if (counter == 2)
                        {
                            pictureBox3.Image = Image.FromStream(stream);
                            pictureBox3.SizeMode = PictureBoxSizeMode.StretchImage;
                            img3SaveIndex = index;
                        }

                        if (counter < 2)
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
            if (index >= indexLimit)
            {
                MessageBox.Show("No More Images!", "No More Images", MessageBoxButtons.OK);
            }
            else
            {
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
                    for (int i = index; i < index+3; i++)
                    {
                        imageUrl = imageList[i];

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
                                img1SaveIndex = index; //Keeping track of image index if I need to download it from the list later
                            }
                            else if (counter == 1)
                            {
                                pictureBox2.Image = Image.FromStream(stream);
                                pictureBox2.SizeMode = PictureBoxSizeMode.StretchImage;
                                img2SaveIndex = index + 1;
                            }
                            else if (counter == 2)
                            {
                                pictureBox3.Image = Image.FromStream(stream);
                                pictureBox3.SizeMode = PictureBoxSizeMode.StretchImage;
                                img3SaveIndex = index + 2;
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
                                string imageUrl = imageList[img3SaveIndex];

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
        private void pictureBox2_DoubleClick(object sender, EventArgs e)
        {
            string message = "Do you want to save the selected image?";
            string title = "Save Image";
            MessageBoxButtons buttons = MessageBoxButtons.YesNo;
            DialogResult result = MessageBox.Show(message, title, buttons);

            SaveFileDialog save = new SaveFileDialog();
            save.Filter = "JPeg Image|*.jpg";
            save.Title = "Save an Image File";

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
                                string imageUrl = imageList[img2SaveIndex];

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
        private void pictureBox1_DoubleClick(object sender, EventArgs e)
        {
            string message = "Do you want to save the selected image?";
            string title = "Save Image";
            MessageBoxButtons buttons = MessageBoxButtons.YesNo;
            DialogResult result = MessageBox.Show(message, title, buttons);

            SaveFileDialog save = new SaveFileDialog();
            save.Filter = "JPeg Image|*.jpg";
            save.Title = "Save an Image File";

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
                                string imageUrl = imageList[img1SaveIndex];

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

        private void makeSlideButton_Click(object sender, EventArgs e)
        {
            DialogResult result = MessageBox.Show("Would You Like To Create a PowerPoint?", "PowerPoint Creation", MessageBoxButtons.YesNo);
            if (result == DialogResult.Yes)
                CreatePPTX();
            else
            {
                result = MessageBox.Show("Would You Like To Add A Slide To An Existing PowerPoint?", "Edit PowerPoint", MessageBoxButtons.YesNo);
                if(result == DialogResult.Yes)
                    LoadPPTx();
                else
                {
                    //close
                }
            }
        }

        protected void LoadPPTx()
        {
            var filePath = string.Empty;

            using (OpenFileDialog openFileDialog = new OpenFileDialog())
            {
                openFileDialog.InitialDirectory = "c:\\";
                openFileDialog.Filter = "pptx files (*.pptx)|*.pptx";
                openFileDialog.FilterIndex = 1;
                openFileDialog.RestoreDirectory = true;

                if (openFileDialog.ShowDialog() == DialogResult.OK)
                {
                    //Get the path of PowerPoint file
                    filePath = openFileDialog.FileName;
                }
            }

            FileInfo fileCheck = new FileInfo(filePath); //Gets the file path to the sample pptx that is getting saved
            if (!IsFileLocked(fileCheck)) //If the file is not open then create and save PowerPoint else throw error message
            {
                //Opens an existing PowerPoint presentation.
                IPresentation pptxDoc = Presentation.Open(filePath);

                //Creates a slide at the end of the PowerPoint presentation
                ISlide slide = pptxDoc.Slides.Add(SlideLayoutType.TitleAndContent);

                //Add title content to the slide by accessing the title placeholder of the TitleOnly layout-slide
                IShape titleShape = slide.Shapes[0] as IShape;
                titleShape.TextBody.AddParagraph(pptTitle).HorizontalAlignment = HorizontalAlignmentType.Center;

                //Adds content to the text box
                IShape descriptionShape = slide.Shapes[1] as IShape;
                descriptionShape.TextBody.AddParagraph(pptText);

                //Saves the Presentation to the file system.
                pptxDoc.Save("Output.pptx");

                //Following algorithm gets all the images that were selected and puts them into the pptx
                if (imageToPptIndex.Count != 0)
                {
                    double imageOffset = 499.79;

                    using (WebClient webClient = new WebClient())
                    {
                        for (int i = 0; i < imageToPptIndex.Count; i++)
                        {
                            string imageUrl = imageList[(int)imageToPptIndex[i]];

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

                            //Gets a picture as stream.
                            MemoryStream stream = new MemoryStream(imageBytes);

                            //Adds the picture to a slide by specifying its size and position.
                            slide.Shapes.AddPicture(stream, imageOffset + i * 10, 238.59, 364.54, 192.16); //Offset each image so the user can see the different images on the pptx
                            stream.Close();
                        }
                    }
                }

                pptxDoc.Save(filePath); //Saves the pptx
                pptxDoc.Close();    //closes pptx stream

                //Dialog box opens, asking the user if they want to open the pptx that was created. If yes then the pptx opens if not then the box closes
                DialogResult result = MessageBox.Show("Slide has been added! Would you like to open the PowerPoint?", "Slide Added", MessageBoxButtons.YesNo);
                if (result == DialogResult.Yes)
                {
                    Process proc = Process.Start(filePath); //Opens and runs the pptx
                }
                else
                {
                    //close
                }
            }
            else
            {
                MessageBox.Show("File is already open and can not be saved!\nPlease close The PowerPoint", "PowerPoint Open", MessageBoxButtons.OK);
            }
        }

        protected void CreatePPTX()
        {
            var filePath = string.Empty;

            MessageBox.Show("Please Select The Folder You Would Like The PowerPoint To Be In", "Folder Select", MessageBoxButtons.OK);

            using (var fbd = new FolderBrowserDialog())
            {
                DialogResult fileResult = fbd.ShowDialog();

                if (fileResult == DialogResult.OK && !string.IsNullOrWhiteSpace(fbd.SelectedPath))
                {
                    filePath = fbd.SelectedPath + "\\Sample.pptx";
                }
            }
            if (filePath != "")
            {
                if (IsSampleFileMade(filePath)) //If the file is made and is not open then create and save PowerPoint else throw error message
                {
                    DialogResult overwriteResult = MessageBox.Show("A PowerPoint With The Name Sample.pptx Is Already Made\nWould You Like To Save Over It?", "PowerPoint Already Exists", MessageBoxButtons.YesNo);
                    if (overwriteResult == DialogResult.Yes)
                    {
                        IPresentation pptxDoc = Presentation.Create(); //Creates new powerpoint

                        ISlide slide = pptxDoc.Slides.Add(SlideLayoutType.TitleAndContent); //Creates new slide

                        //Add title content to the slide by accessing the title placeholder of the TitleOnly layout-slide
                        IShape titleShape = slide.Shapes[0] as IShape;
                        titleShape.TextBody.AddParagraph(pptTitle).HorizontalAlignment = HorizontalAlignmentType.Center;

                        //Adds content to the text box
                        IShape descriptionShape = slide.Shapes[1] as IShape;
                        descriptionShape.TextBody.AddParagraph(pptText);

                        //Following algorithm gets all the images that were selected and puts them into the pptx
                        if (imageToPptIndex.Count != 0)
                        {
                            double imageOffset = 499.79;

                            using (WebClient webClient = new WebClient())
                            {
                                for (int i = 0; i < imageToPptIndex.Count; i++)
                                {
                                    string imageUrl = imageList[(int)imageToPptIndex[i]];

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

                                    //Gets a picture as stream.
                                    MemoryStream stream = new MemoryStream(imageBytes);

                                    //Adds the picture to a slide by specifying its size and position.
                                    slide.Shapes.AddPicture(stream, imageOffset + i * 10, 238.59, 364.54, 192.16); //Offset each image so the user can see the different images on the pptx
                                    stream.Close();
                                }
                            }
                        }

                        pptxDoc.Save(filePath); //Saves the pptx
                        pptxDoc.Close();    //closes pptx stream

                        //Dialog box opens, asking the user if they want to open the pptx that was created. If yes then the pptx opens if not then the box closes
                        DialogResult result = MessageBox.Show("PowerPoint has been created! Would you like to open it?", "PowerPoint Created", MessageBoxButtons.YesNo);
                        if (result == DialogResult.Yes)
                        {
                            Process proc = Process.Start(filePath); //Opens and runs the pptx
                        }
                        else
                        {
                            //close
                        }
                    }
                    else
                    {
                        //close
                    }
                }
                else
                {
                    if (File.Exists(filePath)) //If file is made and the file exists then the pptx must be open. Throw error message
                    {
                        MessageBox.Show("File is already open and can not be saved!\nPlease close The PowerPoint", "PowerPoint Open", MessageBoxButtons.OK);
                    }
                    else //File is not made yet
                    {
                        IPresentation pptxDoc = Presentation.Create(); //Creates new powerpoint

                        ISlide slide = pptxDoc.Slides.Add(SlideLayoutType.TitleAndContent); //Creates new slide

                        //Add title content to the slide by accessing the title placeholder of the TitleOnly layout-slide
                        IShape titleShape = slide.Shapes[0] as IShape;
                        titleShape.TextBody.AddParagraph(pptTitle).HorizontalAlignment = HorizontalAlignmentType.Center;

                        //Adds content to the text box
                        IShape descriptionShape = slide.Shapes[1] as IShape;
                        descriptionShape.TextBody.AddParagraph(pptText);

                        //Following algorithm gets all the images that were selected and puts them into the pptx
                        if (imageToPptIndex.Count != 0)
                        {
                            double imageOffset = 499.79;

                            using (WebClient webClient = new WebClient())
                            {
                                for (int i = 0; i < imageToPptIndex.Count; i++)
                                {
                                    string imageUrl = imageList[(int)imageToPptIndex[i]];

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

                                    //Gets a picture as stream.
                                    MemoryStream stream = new MemoryStream(imageBytes);

                                    //Adds the picture to a slide by specifying its size and position.
                                    slide.Shapes.AddPicture(stream, imageOffset + i * 10, 238.59, 364.54, 192.16); //Offset each image so the user can see the different images on the pptx
                                    stream.Close();
                                }
                            }
                        }

                        pptxDoc.Save(filePath); //Saves the pptx
                        pptxDoc.Close();    //closes pptx stream

                        //Dialog box opens, asking the user if they want to open the pptx that was created. If yes then the pptx opens if not then the box closes
                        DialogResult result = MessageBox.Show("PowerPoint has been created! Would you like to open it?", "PowerPoint Created", MessageBoxButtons.YesNo);
                        if (result == DialogResult.Yes)
                        {
                            Process proc = Process.Start(filePath); //Opens and runs the pptx
                        }
                        else
                        {
                            //close
                        }
                    }
                }
            }
            else
            {
                //close
            }
        }

        protected virtual bool IsFileLocked(FileInfo file)
        {
            try
            {
                using (FileStream stream = file.Open(FileMode.Open, FileAccess.Read, FileShare.None))
                {
                    stream.Close();
                }
            }
            catch (IOException)
            {
                //the file is unavailable because it is:
                //still being written to
                //or being processed by another thread
                //or does not exist (has already been processed)
                return true;
            }

            //file is not locked
            return false;
        }
        protected virtual bool IsSampleFileMade(string file)
        {
            string path = file;
            try
            {
                using (FileStream stream = File.Open(path, FileMode.Open))
                {
                    stream.Close();
                }
            }
            catch (IOException)
            {
                //the has not been made
                return false;
            }

            //file is made
            return true;
        }

        private void pictureBox3_Click(object sender, EventArgs e)
        {
            //Message box initialization
            string message = "Do you want to put the selected image in the PowerPoint?";
            string title = "Copy Image";
            MessageBoxButtons buttons = MessageBoxButtons.YesNo;
            DialogResult result = MessageBox.Show(message, title, buttons);

            if (result==DialogResult.Yes)
            {
                imageToPptIndex.Add(img3SaveIndex); //adds the image index in picture box 3 an array of images to copy to the PowerPoint
            }
            else
            {
                //close
            }
        }

        private void pictureBox2_Click(object sender, EventArgs e)
        {
            //Message box initialization
            string message = "Do you want to put the selected image in the PowerPoint?";
            string title = "Copy Image";
            MessageBoxButtons buttons = MessageBoxButtons.YesNo;
            DialogResult result = MessageBox.Show(message, title, buttons);

            if (result == DialogResult.Yes)
            {
                imageToPptIndex.Add(img2SaveIndex); //adds the image index in picture box 2 an array of images to copy to the PowerPoint
            }
            else
            {
                //close
            }
        }

        private void pictureBox1_Click(object sender, EventArgs e)
        {
            //Message box initialization
            string message = "Do you want to put the selected image in the PowerPoint?";
            string title = "Copy Image";
            MessageBoxButtons buttons = MessageBoxButtons.YesNo;
            DialogResult result = MessageBox.Show(message, title, buttons);

            if (result == DialogResult.Yes)
            {
                imageToPptIndex.Add(img1SaveIndex); //adds the image index in picture box 1 an array of images to copy to the PowerPoint
            }
            else
            {
                //close
            }
        }

        private void clearPicturesButton_Click(object sender, EventArgs e)
        {
            //Message box initialization
            string message = "Do you want to clear you selected images?";
            string title = "Clear Images";
            MessageBoxButtons buttons = MessageBoxButtons.YesNo;
            DialogResult result = MessageBox.Show(message, title, buttons);

            if (result == DialogResult.Yes)
            {
                imageToPptIndex.Clear();
                MessageBox.Show("Your selected images have been cleared", "Cleared Images", MessageBoxButtons.OK);
            }
            else
            {
                //close
            }
        }
    }
}
