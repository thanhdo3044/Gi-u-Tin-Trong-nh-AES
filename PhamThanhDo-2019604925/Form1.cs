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
using Aspose.Words;
using GiauTinTrongAnh;

namespace PhamThanhDo_2019604925
{
	public partial class Form1 : Form
	{
		public Form1()
		{
			InitializeComponent();
		}

        static string chuoiGiaiMa;

        static string chuoiMaHoa;

        static Bitmap anhBanDau;

        static Bitmap anhDaGiauTin;

        static string url1, url2;

        static string urlGoc;

        static string pathResultImage;


        // Giấu thông tin vào ảnh
        public static Bitmap EncryptImage(string data, string password)
        {
            chuoiMaHoa = AES.Encrypt(data, password);    // Mã hoá thông tin

            anhBanDau = SteganographyHelper.CreateNonIndexedImage(Image.FromFile(urlGoc));

            anhDaGiauTin = SteganographyHelper.MergeText(chuoiMaHoa, anhBanDau);

            //imageWithHiddenData.Save(pathResultImage);
            //imageWithHiddenData.Save(@"D:\images\Harry_encrypt1.png");
            return anhDaGiauTin;

        }
        // Tách thông tin từ ảnh
        public static string DecryptImage(string _PASSWORD, Bitmap image)
        {
            //string _PASSWORD = "password";
            //string pathImageWithHiddenInformation = @"D:\images\Harry_encrypt.png";

            string encryptedData = SteganographyHelper.ExtractText(image);

            chuoiGiaiMa = AES.Decrypt(encryptedData, _PASSWORD);// Giải mã thông tin
            return chuoiGiaiMa;

        }

        private void pictureBefor_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog1 = new OpenFileDialog();
            openFileDialog1.Title = "Choose an Image";
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                url1 = openFileDialog1.FileName; // get pathFile ảnh dùng để giấu thông tin 
                Bitmap oldBitmap = new Bitmap(url1);
                URLGiauTin.Text = url1;
                AnhTruocGiau.Image = oldBitmap;
            }
            else
            {
                //inPath2 = "";
            }
        }

        private void Tai_Click(object sender, EventArgs e)
        {
            OpenFileDialog open = new OpenFileDialog();
            open.Filter = "txt files (.txt)|.txt|All files (.)|*.*";
            if (open.ShowDialog() == DialogResult.OK)
            {
                StreamReader reader = new StreamReader(open.FileName);
                NoiDungCanGiau.Text = reader.ReadToEnd();

            }
        }

        private void GiauTin_Click(object sender, EventArgs e)
        {
            try
            {
                urlGoc = URLGiauTin.Text;
                if (string.IsNullOrEmpty(urlGoc))
                {
                    MessageBox.Show("Vui lòng chọn ảnh trước khi giấu tin !");
                    return;
                }
                pathResultImage = URLSave.Text;
                Bitmap newBitmap = EncryptImage(NoiDungCanGiau.Text, txtkhoaTachTin.Text); ;
                //txtpathFileOld.Text = inPath2;
                AnhSauKhiGiau.Image = newBitmap;
                MessageBox.Show("Giấu tin thành công !");
            }
            catch (Exception ex)
            {
                MessageBox.Show("Giấu tin thất bại !" + ex);
            }
        }

        private void LamLai_Click(object sender, EventArgs e)
        {
            NoiDungCanGiau.Clear();
            txtkhoaTachTin.Clear();
            AnhTruocGiau.Image = null;
            AnhSauKhiGiau.Image = null;
            URLGiauTin.Clear();
            URLSave.Clear();
        }

        private void GiaiMa_Click(object sender, EventArgs e)
        {
            try
            {
                if (string.IsNullOrEmpty(KhoaGiaiMa.Text))
                {
                    MessageBox.Show("Vui lòng nhập mật khẩu giải mã !");
                    return;
                }
                if (string.IsNullOrEmpty(URLAnhGiaiMa.Text))
                {
                    MessageBox.Show("Vui lòng chọn ảnh trước khi giải mã !");
                    return;
                }

                //var doc = new Document(NoiDungGiaiMa.Text);
                //NoiDungGiaiMa.Text = DecryptImage(NoiDungGiaiMa.Text, AnhGiaiMa.Image as Bitmap);

                //NoiDungGiaiMa.Text = DecryptImage(dox, AnhGiaiMa.Image as Bitmap);

                //khai bao doc


                //Image.FromFile(pathOriginalImage);
                NoiDungGiaiMa.Text = DecryptImage(KhoaGiaiMa.Text, AnhGiaiMa.Image as Bitmap);

                ////su ly hien thi trong noi dung
                MessageBox.Show("Giải mã thành công !");
            }
            catch (Exception ex)
            {
                MessageBox.Show("Giải mã thất bại !" + ex);
            }
        }

        private void ResetTachTin_Click(object sender, EventArgs e)
        {
            NoiDungGiaiMa.Clear();
            AnhGiaiMa.Image = null;
            URLAnhGiaiMa.Clear();
            KhoaGiaiMa.Clear();
        }

        private void MoAnhTachTin_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog1 = new OpenFileDialog();
            openFileDialog1.Title = "Choose an Image ";
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                url1 = openFileDialog1.FileName; // get pathFile ảnh dùng để giấu thông tin 
                Bitmap oldBitmap = new Bitmap(url1);
                URLAnhGiaiMa.Text = url1;
                AnhGiaiMa.Image = oldBitmap;
            }
            else
            {
                //inPath1 = "";
            }
        }

        private void MoAnh_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog1 = new OpenFileDialog();
            openFileDialog1.Title = "Choose an Image";
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                url1 = openFileDialog1.FileName; // get pathFile ảnh dùng để giấu thông tin 
                Bitmap oldBitmap = new Bitmap(url1);
                URLGiauTin.Text = url1;
                AnhTruocGiau.Image = oldBitmap;
            }
            else
            {
                //inPath1 = "";
            }
        }

        private void Moifiletxt_dox_Click(object sender, EventArgs e)
        {
            using (OpenFileDialog mofiledox = new OpenFileDialog() { ValidateNames = true, Multiselect = false, Filter = "|*.docx;*.txt" })
            {
                if (mofiledox.ShowDialog() == DialogResult.OK)
                {
                    object readOnly = false;
                    object visible = true;
                    object save = false;
                    object fileName = mofiledox.FileName;
                    object newTemplate = false;
                    object doxType = 0;
                    object missing = Type.Missing;
                    Microsoft.Office.Interop.Word._Document document;
                    Microsoft.Office.Interop.Word._Application application = new Microsoft.Office.Interop.Word.Application() { Visible = false };
                    document = application.Documents.Open(ref fileName, ref missing, ref readOnly, ref missing, ref missing, ref missing, ref missing,
                        ref missing, ref missing, ref missing, ref missing, ref visible, ref missing, ref missing, ref missing, ref missing);
                    document.ActiveWindow.Selection.WholeStory();
                    document.ActiveWindow.Selection.Copy();
                    IDataObject dataObject = Clipboard.GetDataObject();
                    NoiDungCanGiau.Rtf = dataObject.GetData(DataFormats.Rtf).ToString();
                    application.Quit(ref missing, ref missing, ref missing);
                }
            }
            // mo file txt
            /*
            OpenFileDialog mofile = new OpenFileDialog();
            mofile.Filter = "|*.txt";
            if (mofile.ShowDialog() == DialogResult.OK)
			{
                StreamReader read = new StreamReader(mofile.FileName);
                NoiDungCanGiau.Text = read.ReadToEnd();
                read.Close();
            }
            */

        }

        private void label3_Click(object sender, EventArgs e)
        {

        }

        private void button1_Click(object sender, EventArgs e)
        {
            OpenFileDialog mofile = new OpenFileDialog();
            mofile.Filter = "|*.txt;*.docx";
            if (mofile.ShowDialog() == DialogResult.OK)
            {
                StreamReader read = new StreamReader(mofile.FileName);
                txtkhoaTachTin.Text = read.ReadToEnd();
                read.Close();
            }
        }

        private void btnMofile_TackTin_Click(object sender, EventArgs e)
        {
            OpenFileDialog mofile = new OpenFileDialog();
            mofile.Filter = "|*.txt;*.docx";
            if (mofile.ShowDialog() == DialogResult.OK)
            {
                StreamReader read = new StreamReader(mofile.FileName);
                KhoaGiaiMa.Text = read.ReadToEnd();
                read.Close();
            }
        }

        private void btn_LuuFile_TachTin_Click(object sender, EventArgs e)
        {
            SaveFileDialog luufile = new SaveFileDialog();
            luufile.Filter = "|*.txt;*.docx";
            if (luufile.ShowDialog() == DialogResult.OK)
            {
                if (luufile.Filter == "|*.txt")
                {
                    StreamWriter writer = new StreamWriter(luufile.FileName);
                    writer.WriteLine(NoiDungGiaiMa.Text);
                    writer.Close();
                    MessageBox.Show("Lưu dữ liệu thành công !");
                }
                if (luufile.Filter == "|*.docx")
                {
                    var doc = new Document(NoiDungGiaiMa.Text);
                    var builder = new DocumentBuilder(doc);

                    // Chèn văn bản vào đầu tài liệu.
                    builder.MoveToDocumentStart();
                    builder.Write(NoiDungGiaiMa.Text);
                    doc.Save("NoiDung_Luu.docx");

                    MessageBox.Show("Lưu thành công !");
                }
            }

        }

        private void Save_Click(object sender, EventArgs e)
        {
            if (AnhSauKhiGiau.Image == null)
            {
                MessageBox.Show("Vui lòng giấu tin trước khi lưu ảnh !");
                return;
            }
            SaveFileDialog saveFileDialog = new SaveFileDialog();

            saveFileDialog.Title = "Save Image";

            saveFileDialog.Filter = "Image Files (*.png;*.jpg; *.bmp)|*.png;*.jpg;*.bmp" + "|" + "All Files (*.*)|*.*";
            //openFileDialog1.Filter = "txt files (*.jpg)|*.jpg|All files (*.*)|*.*";
            //SaveFileDialog.Filter = "Image files (*.jpg, *.jpeg, *.jpe, *.jfif, *.png) | *.jpg; *.jpeg; *.jpe; *.jfif; *.png|All files (*.*)|*.*";
            if (saveFileDialog.ShowDialog() == DialogResult.OK)
            {
                url2 = saveFileDialog.FileName; // get pathFile ảnh đã giấu tin

                AnhSauKhiGiau.Image.Save(url2);
                MessageBox.Show("Lưu ảnh thành công");
            }
            else
            {
                url2 = "";
            }
        }

    }
}
