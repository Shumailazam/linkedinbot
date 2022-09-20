using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using excel = Microsoft.Office.Interop.Excel;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using System.Text.RegularExpressions;

namespace WindowsFormsApp1
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            //string path = @"E:\SpecFlow\linkedin\WindowsFormsApp1\sh.csv";
            //Application excel = new Application();
            //Workbook wb = excel.Workbooks.Open(path);
            //Worksheet excelSheet = wb.ActiveSheet;

            //Read the first cell
            //string test = excelSheet.Cells[1, 1].Value.ToString();


            excel.Application xlapp = new excel.Application();
            excel.Workbook x1workbook = xlapp.Workbooks.Open(@"E:\SpecFlow\linkedin\WindowsFormsApp1\sh.csv");
            excel._Worksheet x1worksheet = x1workbook.Sheets[1];

            excel.Range x1range = x1worksheet.UsedRange;

            string website;



            //if(checkBox1.Checked == true)
            //{
            //    //website = Convert.ToString(x1range.Cells[4][j].value2);

            //    //string y = website.ToString();

            //    IWebDriver driver = new ChromeDriver();
            //    driver.Navigate().GoToUrl("https://www.linkedin.com");
            //    System.Threading.Thread.Sleep(4000);

            //    driver.FindElement(By.XPath("//*[@class='sign-in-form-container']/div/form")).Click();
            //    System.Threading.Thread.Sleep(4000);

            //    driver.SwitchTo().Window(driver.WindowHandles[1]);


            //    driver.FindElement(By.Id("identifierId")).SendKeys("meenupatelmp");
            //    System.Threading.Thread.Sleep(4000);

            //    driver.FindElement(By.XPath("//*[@id='identifierNext']/div/button/span")).Click();
            //    System.Threading.Thread.Sleep(4000);
            //}

            //else
            //{
            IWebDriver driver = new ChromeDriver();
            driver.Navigate().GoToUrl("https://www.linkedin.com");
            driver.Manage().Window.Maximize();
            System.Threading.Thread.Sleep(4000);
            driver.FindElement(By.Id("session_key")).SendKeys("menupatelmp@gmail.com");
            System.Threading.Thread.Sleep(4000);
            driver.FindElement(By.Id("session_password")).SendKeys("OntologY1!");
            System.Threading.Thread.Sleep(4000);
            driver.FindElement(By.XPath("//*[@class='sign-in-form-container']/form/button")).Click();
            System.Threading.Thread.Sleep(5000);
            for (int j = 2; j < 100; j++)
            {

                website = Convert.ToString(x1range.Cells[4][j].value2);

                string y = website.ToString();
                driver.Navigate().GoToUrl(website);
                System.Threading.Thread.Sleep(5000);
                try
                {
                    IWebElement texterr = driver.FindElement(By.XPath("//*[@id='ember10']/h2[1]"));
                    string error = texterr.Text;
                    System.Threading.Thread.Sleep(5000);
                    if (error != "" || error != null)
                    {

                    }
                    else
                    {
                        driver.FindElement(By.XPath("//*[@id='ember10']/a")).Click();
                        System.Threading.Thread.Sleep(4000);
                    }
                    System.Threading.Thread.Sleep(5000);
                }
                catch (Exception ex)
                {

                    string ref1;
                    try
                    {
                        //getting name first 

                        try
                        {
                            IWebElement name = driver.FindElement(By.XPath("//*[@class='ph5 pb5']/div[2]/div/div"));

                            var n1 = name.Text;

                            String[] namearr = n1.Split(' ');
                            var fname = namearr[0];
                            var fnam1 = namearr[1];
                            int jk = fname.Length;
                            int p = fnam1.Length;
                            if(jk>=3)
                            {
                                IWebElement checktxt = driver.FindElement(By.XPath("//*[@class='ph5 pb5']/div[3]/div/div/button"));
                                var nameer = checktxt.Text;
                                if (checktxt.Text.Contains("Follow"))
                                {
                                    try
                                    {
                                        driver.FindElement(By.XPath("//*[@class='pvs-profile-actions ']/div[3]")).Click();
                                        System.Threading.Thread.Sleep(4000);
                                        driver.FindElement(By.XPath("//*[@class='ph5 pb5']/div[3]/div/div[3]/div/div/ul/li[4]/div/span[1]")).Click();

                                        System.Threading.Thread.Sleep(4000);


                                        IWebElement verifytxt = driver.FindElement(By.XPath("//*[@id='artdeco-modal-outlet']/div/div/div[1]/h2"));

                                        System.Threading.Thread.Sleep(3000);
                                        if (verifytxt.Text == "You can customize this invitation")
                                        {

                                            System.Threading.Thread.Sleep(3000);
                                            driver.FindElement(By.XPath("//*[@id='artdeco-modal-outlet']/div/div/div[3]/button[1]/span")).Click();

                                            driver.FindElement(By.Id("custom-message")).SendKeys("Hello " + fname + "");

                                            System.Threading.Thread.Sleep(4000);
                                            driver.FindElement(By.XPath("//*[@id='artdeco-modal-outlet']/div/div/div[3]/button[2]")).Click();
                                            //System.Threading.Thread.Sleep(3000);
                                            //msgtxt.SendKeys("Hello "+fname+"");
                                            //System.Threading.Thread.Sleep(3000);
                                        }
                                    }
                                    catch(Exception ex21)
                                    {

                                    }
                                   
                                }
                                else
                                {
                                    //connect code
                                    driver.FindElement(By.XPath("//*[@class='artdeco-dropdown__content-inner']/ul/li[4]/div")).Click();
                                }
                            }
                            else if (p>=3)
                            {
                                IWebElement checktxt = driver.FindElement(By.XPath("//*[@class='ph5 pb5']/div[3]/div/div/button"));
                                var nameer = checktxt.Text;
                                if(checktxt.Text.Contains("Follow"))
                                {
                                    driver.FindElement(By.XPath("//*[@class='pvs-profile-actions ']/div[3]")).Click();

                                    driver.FindElement(By.XPath("//*[@class='artdeco-dropdown__content-inner']/ul/li[4]/div")).Click();
                                }
                                else
                                {
                                    //connect code
                                    driver.FindElement(By.XPath("//*[@class='artdeco-dropdown__content-inner']/ul/li[4]/div")).Click();
                                }
                            }

                        }

                        catch(Exception ex2)
                        {
                            //first name error exception
                        }

                        
                        IWebElement follow = driver.FindElement(By.XPath("//*[@class='entry-point']/a"));

                        ref1 = follow.GetAttribute("href");
                        driver.Navigate().GoToUrl(ref1);
                        //follow.Click();
                    }
                    catch(Exception ex1)
                    {

                    }
                }
                
            }
        }

        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {

        }

        private void button2_Click(object sender, EventArgs e)
        {
            IWebDriver driver = new ChromeDriver();
            excel.Workbook x1workbook = null;
            button2.Enabled = false;
            if (!(String.IsNullOrWhiteSpace(textBox1.Text)))
            {
                try
                {
                    excel.Application xlapp = new excel.Application();
                    var path = textBox3.Text.Trim();
                    x1workbook = xlapp.Workbooks.Open(@path);
                    //excel.Workbook x1workbook = xlapp.Workbooks.Open(@"E:\SpecFlow\linkedin\WindowsFormsApp1\sh.csv");
                    excel._Worksheet x1worksheet = x1workbook.Sheets[1];
                    excel.Range x1range = x1worksheet.UsedRange;
                    
                    string website;


                    //chrome driver start

                    
                    driver.Navigate().GoToUrl("https://www.linkedin.com");
                    driver.Manage().Window.Maximize();
                    System.Threading.Thread.Sleep(4000);
                    driver.FindElement(By.Id("session_key")).SendKeys(textBox1.Text.Trim());
                    System.Threading.Thread.Sleep(4000);
                    driver.FindElement(By.Id("session_password")).SendKeys(textBox2.Text.Trim());
                    System.Threading.Thread.Sleep(4000);
                    driver.FindElement(By.XPath("//*[@class='sign-in-form-container']/form/button")).Click();
                    System.Threading.Thread.Sleep(5000);

                    try
                    {
                        IWebElement invalidpass = driver.FindElement(By.Id("error-for-password"));


                        if (invalidpass.Text == "That's not the right password. Try again or ")
                        {
                            MessageBox.Show("Invalid Id or password display!!!!");
                            driver.Quit();
                            button2.Enabled = true;
                            textBox1.Clear();
                            textBox2.Clear();
                            x1workbook.Close();
                        }
                        else
                        {
                            driver.Quit();
                            button2.Enabled = true;
                            textBox1.Clear();
                            textBox2.Clear();
                            x1workbook.Close();
                        }
                    }
                    catch (Exception ex)
                    {
                        //ID AND PASSWORD IS VALID
                        try
                        {
                            IWebElement ip = driver.FindElement(By.XPath("//*[@class='sign-in-form-container']/form/div/div/p"));
                            System.Threading.Thread.Sleep(4000);
                            if (ip.Text == "Please enter a valid email address or mobile number.")
                            {
                                MessageBox.Show("Invalid Id or password display!!!!");
                                driver.Quit();
                                button2.Enabled = true;
                                textBox1.Clear();
                                textBox2.Clear();
                                x1workbook.Close();
                            }
                        }
                        catch (Exception ex1)
                        {
                            //Valid Login Code successfully logged in 
                            int jk = Int32.Parse(textBox4.Text.Trim());
                            System.Threading.Thread.Sleep(4000);
                            for (int j = 2; j <= jk ; j++)
                            {
                                System.Threading.Thread.Sleep(4000);
                                website = Convert.ToString(x1range.Cells[4][j].value2);

                                string y = website.ToString();
                                driver.Navigate().GoToUrl(website);
                                System.Threading.Thread.Sleep(5000);


                                //Valid link or not 
                                try
                                {
                                    IWebElement texterr = driver.FindElement(By.XPath("//*[@id='ember10']/h2[1]"));
                                    string error = null;
                                    error = texterr.Text;
                                    System.Threading.Thread.Sleep(5000);
                                    if (error == "This page doesn’t exist" || error != null || error !="")
                                    {
                                        // error occured link not found

                                        continue;
                                    }
                                    else

                                    {
                                        //driver.FindElement(By.XPath("//*[@id='ember10']/a")).Click();
                                        //System.Threading.Thread.Sleep(4000);
                                    }
                                    System.Threading.Thread.Sleep(5000);
                                }
                                catch (Exception ex2)
                                {
                                    string ref1;
                                    System.Threading.Thread.Sleep(4000);
                                    IWebElement name = driver.FindElement(By.XPath("//*[@class='ph5 pb5']/div[2]/div/div"));

                                    var n1 = name.Text;

                                    String[] namearr = n1.Split(' ');
                                    var fname = namearr[0];
                                    var fnam1 = namearr[1];
                                    int jkp = fname.Length;
                                    int p = fnam1.Length;

                                    // name checking emoji exist or not 
                                    if (jkp < 3)
                                    {
                                        //emopji is there in name
                                    }
                                    else
                                    {
                                        // check if link is alraedy followed
                                        System.Threading.Thread.Sleep(4000);
                                        IWebElement checktxt = driver.FindElement(By.XPath("//*[@class='ph5 pb5']/div[3]/div/div/button"));
                                        //below contain more in already following account
                                        var nameer = checktxt.Text;

                                        try
                                        {
                                            //checking element to identify connect or not
                                            IWebElement connectreq = driver.FindElement(By.XPath("//*[@class='ph5 pb5']/div[3]/div"));
                                            var msg = connectreq.Text;
                                            msg = msg.Replace("\r\n", " ");
                                            var v1 = msg.Replace("\r\n", " ");
                                            String[] options = msg.Split(' ');

                                            for (int h = 0; h < options.Length; h++)
                                            {
                                                if (options[h] == "Follow")
                                                {
                                                    //follow click
                                                }
                                                else if (options[h] == "Following")
                                                {
                                                    continue;
                                                }
                                                else if (options[h] == "Pending")
                                                {
                                                    break;
                                                }

                                                else if (options[h] == "Connect")
                                                {
                                                    IWebElement vert = driver.FindElement(By.XPath("//*[@class='ph5 pb5']/div[3]/div[1]/div/button/span"));
                                                    string verttext = vert.Text;
                                                    if (verttext == "Connect")
                                                    {
                                                        driver.FindElement(By.XPath("//*[@class='ph5 pb5']/div[3]/div[1]/div/button/span")).Click();
                                                        System.Threading.Thread.Sleep(4000);
                                                        IWebElement verifytxt = driver.FindElement(By.XPath("//*[@id='artdeco-modal-outlet']/div/div/div[1]/h2"));
                                                        System.Threading.Thread.Sleep(3000);
                                                        if (verifytxt.Text == "You can customize this invitation")
                                                        {
                                                            System.Threading.Thread.Sleep(3000);
                                                            driver.FindElement(By.XPath("//*[@id='artdeco-modal-outlet']/div/div/div[3]/button[1]/span")).Click();
                                                            driver.FindElement(By.Id("custom-message")).SendKeys("Hello " + fname + "");
                                                            System.Threading.Thread.Sleep(4000);
                                                            //send button
                                                            //driver.FindElement(By.XPath("//*[@id='artdeco-modal-outlet']/div/div/div[3]/button[2]")).Click();
                                                            break;
                                                        }
                                                    }
                                                }
                                                else if (options[h] == "More")
                                                {
                                                    try
                                                    {
                                                        //checking what button is available
                                                        IWebElement whattext = driver.FindElement(By.XPath("//*[@class='pvs-profile-actions ']/div[2]"));
                                                        var textcheck = whattext.Text;
                                                        if (textcheck == "Message")
                                                        {
                                                            driver.FindElement(By.XPath("//*[@class='ph5 pb5']/div[3]/div[1]/div[3]")).Click();
                                                            IWebElement moreelementfollow = driver.FindElement(By.XPath("//*[@class='ph5 pb5']/div[3]/div[1]/div[3]/div[1]/div/ul"));
                                                            var morefextra = moreelementfollow.Text;
                                                            morefextra = morefextra.Replace("\r\n", ",");
                                                            String[] m1 = morefextra.Split(',');
                                                            for (int r = 0; r < m1.Length; r++)
                                                            {
                                                                if (m1[r] == "Connect")
                                                                {

                                                                }
                                                            }
                                                        }
                                                        else
                                                        {
                                                            //for skippingalready follow
                                                            driver.FindElement(By.XPath("//*[@class='pvs-profile-actions ']/div[2]")).Click();
                                                            IWebElement moreele = driver.FindElement(By.XPath("//*[@class='ph5 pb5']/div[3]/div[1]/div[2]/div/div/ul"));
                                                            var extra = moreele.Text;
                                                            extra = extra.Replace("\r\n", ",");
                                                            String[] moreextra = extra.Split(',');
                                                            for (int pi = 0; pi < moreextra.Length; pi++)
                                                            {
                                                                //two code for skip and for connection
                                                                if (moreextra[pi] == "Remove Connection")
                                                                {
                                                                    break;
                                                                }
                                                            }
                                                        }


                                                        //already following


                                                    }
                                                    catch (Exception exc)
                                                    {
                                                        driver.FindElement(By.XPath("//*[@class='pvs-profile-actions ']/div[3]")).Click();
                                                        //more option not visible;
                                                    }


                                                }
                                            }
                                        }
                                        catch (Exception exer)
                                        {

                                        }

                                    }
                                        ////below exception occured on already following account
                                        //driver.FindElement(By.XPath("//*[@class='pvs-profile-actions ']/div[3]")).Click();
                                        //System.Threading.Thread.Sleep(4000);
                                        //IWebElement connectverify = driver.FindElement(By.XPath("//*[@class='ph5 pb5']/div[3]/div/div[3]/div/div/ul/li[4]/div/span[1]"));
                                        //var connectverf = connectverify.Text;
                                        //System.Threading.Thread.Sleep(3000);
                                        //if (checktxt.Text.Contains("Follow"))
                                        //{
                                            
                                        //    IWebElement verifytxt = driver.FindElement(By.XPath("//*[@id='artdeco-modal-outlet']/div/div/div[1]/h2"));
                                        //    System.Threading.Thread.Sleep(3000);
                                        //    if (verifytxt.Text == "You can customize this invitation")
                                        //    {
                                        //        //below code is for connection request and message
                                        //        //System.Threading.Thread.Sleep(3000);
                                        //        //driver.FindElement(By.XPath("//*[@id='artdeco-modal-outlet']/div/div/div[3]/button[1]/span")).Click();
                                        //        //driver.FindElement(By.Id("custom-message")).SendKeys("Hello " + fname + "," + textBox5.Text.Trim());
                                        //        //System.Threading.Thread.Sleep(4000);
                                        //        //driver.FindElement(By.XPath("//*[@id='artdeco-modal-outlet']/div/div/div[3]/button[2]")).Click();
                                        //        //System.Threading.Thread.Sleep(3000);
                                                
                                        //        //code for message
                                                
                                                
                                        //    }
                                       // }
                                        //else if (checktxt.Text.Contains("Following"))
                                        //{
                                        //    continue;
                                        //}
                                        //else if(checktxt.Text.Contains("More"))
                                        //{
                                        //    continue;
                                        //}
                                    //}
                                }
                            }
                            System.Threading.Thread.Sleep(4000);
                            //for loop end
                        }
                    }
                    

                }
                catch (Exception ex)
                {
                    button2.Enabled = true;
                    textBox1.Clear();
                    textBox2.Clear();
                    //File exist or not 
                    MessageBox.Show("File Doesn't Exist on the path defined or format is not supported", "Alert!!!!", MessageBoxButtons.OK);
                    x1workbook.Close();
                    
                }

                driver.Quit();
                x1workbook.Close();
                System.Windows.Forms.Application.Exit();
            }
            else
            {
                
                MessageBox.Show("TextBoxes cannot be null or empty!!!!","Alert!!!",MessageBoxButtons.OK);
                textBox1.Clear();
                textBox2.Clear();
                button2.Enabled = true;
                x1workbook.Close();
            }
            
        }

        private void button3_Click(object sender, EventArgs e)
        {
            
        }
    }
}
