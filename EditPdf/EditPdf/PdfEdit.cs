using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Activities;
using System.ComponentModel.Composition;
using System.ComponentModel;
using iTextSharp.text.pdf;
using System.IO;

namespace EditPdf
{
    public class PdfEdit: CodeActivity
    {

        [Category("Input")]
        [RequiredArgument]
        public InArgument<string> InputFilePath { get; set; }

        [Category("Input")]
        [RequiredArgument]
        public InArgument<string> InputFileName { get; set; }

        [Category("Input")]
        [RequiredArgument]
        public InArgument<string> OutputFilePath { get; set; }

        [Category("Input")]
        [RequiredArgument]
        public InArgument<string> OutputFileName { get; set; }

        [Category("Input")]
        [RequiredArgument]
        public InArgument<string> GivenName { get; set; }

        [Category("Input")]
        [RequiredArgument]
        public InArgument<string> FamilyName { get; set; }

        [Category("Input")]
        [RequiredArgument]
        public InArgument<string> Address1 { get; set; }

        [Category("Input")]
        [RequiredArgument]
        public InArgument<string> HouseNo { get; set; }

        [Category("Input")]
        [RequiredArgument]
        public InArgument<string> Address2 { get; set; }

        [Category("Input")]
        [RequiredArgument]
        public InArgument<string> Postcode { get; set; }

        [Category("Input")]
        [RequiredArgument]
        public InArgument<string> City { get; set; }

        [Category("Input")]
        [RequiredArgument]
        public InArgument<string> Country { get; set; }

        [Category("Input")]
        [RequiredArgument]
        public InArgument<string> Gender { get; set; }

        [Category("Input")]
        [RequiredArgument]
        public InArgument<string> Height { get; set; }

        [Category("Input")]
        [RequiredArgument]
        public InArgument<string> DrivingLicense { get; set; }

        [Category("Input")]
        [RequiredArgument]
        public InArgument<string> FavoriteColor { get; set; }


        protected override void Execute(CodeActivityContext context)
        {
            string input_file_path = InputFilePath.Get(context);
            string input_file_name = InputFileName.Get(context);
            string output_file_path = OutputFilePath.Get(context);
            string output_file_name = OutputFileName.Get(context);


            string formFile = input_file_path + "\\" + input_file_name;

            string newFile = output_file_path + "\\" + output_file_name;

            PdfReader reader = new PdfReader(formFile);
            PdfStamper stamper = new PdfStamper(reader, new FileStream(newFile, FileMode.Create));
            AcroFields fields = stamper.AcroFields;

            string givenName = GivenName.Get(context);
            fields.SetField("Given Name Text Box", givenName);

            string familyName = FamilyName.Get(context);
            fields.SetField("Family Name Text Box", familyName);

            string address1 = Address1.Get(context);
            fields.SetField("Address 1 Text Box", address1);

            string houseNo = HouseNo.Get(context);
            fields.SetField("House nr Text Box", houseNo);

            string address2 = Address2.Get(context);
            fields.SetField("Address 2 Text Box", address2);

            string postCode = Postcode.Get(context);
            fields.SetField("Postcode Text Box", postCode);

            string cityField = City.Get(context);
            fields.SetField("City Text Box", cityField);

            string country = Country.Get(context);
            fields.SetField("Country Combo Box", country);

            string gender = Gender.Get(context);
            fields.SetField("Gender List Box", gender);

            string height = Height.Get(context);
            fields.SetField("Height Formatted Field", height);

            string drivingLicense = DrivingLicense.Get(context);
            fields.SetField("Driving License Check Box", drivingLicense);

            string favoriteColor = FavoriteColor.Get(context);
            fields.SetField("Favourite Colour List Box", favoriteColor);

            stamper.Close();
            reader.Close();

        }

        }
    }
