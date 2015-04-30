﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using SACLIENTLib;
using System.IO;
using System.Xml;
using System.Xml.Linq;
using System.Runtime.InteropServices;


namespace ALM_Add_User
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        private SAapi sapi = new SAapi();

        public MainWindow()
        {
            InitializeComponent();
        }

        private void btnBrowse_Click(object sender, RoutedEventArgs e)
        {
            Microsoft.Win32.OpenFileDialog dlg = new Microsoft.Win32.OpenFileDialog();

            // Display OpenFileDialog by calling ShowDialog method 
            Nullable<bool> result = dlg.ShowDialog();


            // Get the selected file name and display in a TextBox 
            if (result == true)
            {
                string filename = dlg.FileName;
                txtFileLocation.Text = filename;
            }
        }

        private void btnLogin_Click(object sender, RoutedEventArgs e)
        {
            //Login and get all domains in XML format
            sapi.Login("http://usman-smappvw04.sncorp.smith-nephew.com:8080/qcbin", txtUsername.Text, txtPassword.Password);
            string xmlDomains = sapi.GetAllDomains();
            XDocument xml = XDocument.Parse(xmlDomains);
            IEnumerable<XElement> allDomains = xml.Root.Elements("TDXItem").Elements("DOMAIN_NAME");

            //For each Xml domain element get just the value
            List<string> finalDomains = new List<string>();
            foreach (XElement domain in allDomains)
            {
                finalDomains.Add(domain.Value);
            }

            //Set dropdown source and enable controls
            drpDomain.ItemsSource = finalDomains;
            drpDomain.IsEnabled = true;
            drpProject.IsEnabled = true;
            txtFileLocation.IsEnabled = true;
            btnBrowse.IsEnabled = true;
            lblLoginStatus.Content = "Logged in as " + txtUsername.Text;

            //disable login fields
            btnLogin.IsEnabled = false;
            txtUsername.IsEnabled = false;
            txtPassword.IsEnabled = false;
        }

        private void drpDomain_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            //When domain selection changes make sure an item is selected before getting all projects for that domain
            if (drpDomain.SelectedIndex >= 0)
            {
                string xmlProjects = sapi.GetAllDomainProjects(drpDomain.SelectedValue.ToString());
                XDocument xml = XDocument.Parse(xmlProjects);
                IEnumerable<XElement> allProjects = xml.Root.Elements("TDXItem").Elements("PROJECT_NAME");

                //Get project values and populate dropdown
                List<string> finalProjects = new List<string>();
                foreach (XElement project in allProjects)
                {
                    finalProjects.Add(project.Value);
                }
                drpProject.ItemsSource = finalProjects;
            }
        }

        private void txtFileLocation_TextChanged(object sender, TextChangedEventArgs e)
        {
            //checks to make sure domain, project, and file have values before enabling add users button
            if (txtFileLocation.Text.Equals(""))
            {
                btnAddUsers.IsEnabled = false;
            }
            else if (drpDomain.SelectedIndex >= 0 && drpProject.SelectedIndex >= 0)
            {
                btnAddUsers.IsEnabled = true;
            }
        }

        private void btnClear_Click(object sender, RoutedEventArgs e)
        {
            //Resets fields
            drpDomain.SelectedIndex = -1;
            drpProject.SelectedIndex = -1;
            txtFileLocation.Text = "";
            lblUploadStatus.Content = "";
        }

        private void btnAddUsers_Click(object sender, RoutedEventArgs e)
        {
            //Open the file and add users and groups to list
            try
            {
                var reader = new StreamReader(File.OpenRead(txtFileLocation.Text));
                List<string> users = new List<string>();
                List<string> groups = new List<string>();
                while (!reader.EndOfStream)
                {
                    var line = reader.ReadLine();
                    var values = line.Split(',');

                    users.Add(values[0]);
                    groups.Add(values[1]);
                }

                //for each user add them to the project and respective group
                int i = 0;
                foreach (string user in users)
                {
                    try
                    {
                        sapi.CreateUserEx(user, "", "", "", "", "", "");
                        sapi.AddUsersToProject(drpDomain.SelectedValue.ToString(), drpProject.SelectedValue.ToString(), user);
                        sapi.AddUsersToGroup(drpDomain.SelectedValue.ToString(), drpProject.SelectedValue.ToString(), groups.ElementAt(i), user);
                    }
                    catch (COMException)
                    {
                        
                    }
                    i++;
                }
                reader.Close();

            }
            catch (Exception)
            {
                MessageBox.Show("Could not open file at " + txtFileLocation.Text);
            }
       
            lblUploadStatus.Content = "Uploading Task Complete!";
        }
    }
}
