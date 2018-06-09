using System;
using System.Diagnostics;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Net.Mail;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace МояПосылка
{
    public partial class FormГлавная : Form
    {
        private DateTime date;

        public FormГлавная()
        {
            InitializeComponent();
            timer.Enabled = true;
            timer.Interval = 8000;
        }

        private void FormГлавная_Load(object sender, EventArgs e)
        {
            // TODO: данная строка кода позволяет загрузить данные в таблицу "мояПосылка4DataSet.ОтправкаEMail". При необходимости она может быть перемещена или удалена.
            this.отправкаEMailTableAdapter.Fill(this.мояПосылка4DataSet.ОтправкаEMail);
            // TODO: данная строка кода позволяет загрузить данные в таблицу "мояПосылка3DataSet.ЛогинПароль". При необходимости она может быть перемещена или удалена.
            this.логинПарольTableAdapter.Fill(this.мояПосылка3DataSet.ЛогинПароль);
            // TODO: данная строка кода позволяет загрузить данные в таблицу "мояПосылка2DataSet.Получение". При необходимости она может быть перемещена или удалена.
            this.получениеTableAdapter.Fill(this.мояПосылка2DataSet.Получение);
            // TODO: данная строка кода позволяет загрузить данные в таблицу "мояПосылкаDataSet.Отправка". При необходимости она может быть перемещена или удалена.
            this.отправкаTableAdapter.Fill(this.мояПосылкаDataSet.Отправка);
            // TODO: данная строка кода позволяет загрузить данные в таблицу "мояПосылкаDataSet.Отправка". При необходимости она может быть перемещена или удалена.
            this.отправкаTableAdapter.Fill(this.мояПосылкаDataSet.Отправка);
            txtШапка.Text = Properties.Settings.Default.txtШапка;
            txtШапка2.Text = Properties.Settings.Default.txtШапка2;
            richТекст.Text = Properties.Settings.Default.richТекст;
            txtДатаЕмайл.Text = DateTime.Now.ToString("dd.MM.yyyy");
            pictureФотоОтправителя.Image = pictureФото.Image;
            labelЖирный.Click += new EventHandler(labelЖирный_firstClick);
            labelРазмер.Click += new EventHandler(labelРазмер_firstClick);
            labelПолужирный.Click += new EventHandler(labelПолужирный_firstClick);
            labelПодчеркнутый.Click += new EventHandler(labelПодчеркнутый_firstClick);
            labelЗачеркнутый.Click += new EventHandler(labelЗачеркнутый_firstClick);
            labelКурсив.Click += new EventHandler(labelКурсив_firstClick);
            labelЦвет.Click += new EventHandler(labelЦвет_firstClick);
            labelЗаголовок.Click += new EventHandler(labelЗаголовок_firstClick);
            labelЛево.Click += new EventHandler(labelЛево_firstClick);
            labelЦентр.Click += new EventHandler(labelЦентр_firstClick);
            labelПраво.Click += new EventHandler(labelПраво_firstClick);
            labelСохранитьКак.Click += new EventHandler(labelСохранитьКак_firstClick);
            dataGridView.Columns[0].Visible = false;
            for (int i = 0; i < advancedDataGridView.RowCount; i++)
            {
                for (int j = 0; j < advancedDataGridView.Columns.Count; j++)
                {

                    switch (advancedDataGridView[j, i].FormattedValue.ToString())
                    {
                        case "Посылка удалена":
                            advancedDataGridView[j, i].Style.BackColor = Color.LightPink;
                            advancedDataGridView[j, i].Style.ForeColor = Color.LightPink;
                            break;
                    }
                }
            }
        }

        private void labelСохранитьКак_firstClick(object sender, EventArgs e)
        {
            labelСохранитьКак.BackColor = Color.SeaGreen;
            panelСохранитьКак.Visible = true;
            labelСохранитьКак.Click -= new EventHandler(labelСохранитьКак_firstClick);
            labelСохранитьКак.Click += new EventHandler(labelСохранитьКак_secondClick);
        }

        private void labelСохранитьКак_secondClick(object sender, EventArgs e)
        {
            labelСохранитьКак.BackColor = Color.DarkSeaGreen;
            panelСохранитьКак.Visible = false;
            labelСохранитьКак.Click += new EventHandler(labelСохранитьКак_firstClick);//включаем первый обработчик
            labelСохранитьКак.Click -= new EventHandler(labelСохранитьКак_secondClick);//отключаем второй обработчик
        }

        private void labelПраво_firstClick(object sender, EventArgs e)
        {
            labelПраво.BackColor = Color.SeaGreen;
            richТекст.SelectionAlignment = HorizontalAlignment.Right;
            labelПраво.Click -= new EventHandler(labelПраво_firstClick);
            labelПраво.Click += new EventHandler(labelПраво_secondClick);
        }

        private void labelПраво_secondClick(object sender, EventArgs e)
        {
            labelПраво.BackColor = Color.DarkSeaGreen;
            richТекст.SelectionAlignment = HorizontalAlignment.Left;
            labelПраво.Click += new EventHandler(labelПраво_firstClick);
            labelПраво.Click -= new EventHandler(labelПраво_secondClick);
        }

        private void labelЦентр_firstClick(object sender, EventArgs e)
        {
            labelЦентр.BackColor = Color.SeaGreen;
            richТекст.SelectionAlignment = HorizontalAlignment.Center;
            labelЦентр.Click -= new EventHandler(labelЦентр_firstClick);
            labelЦентр.Click += new EventHandler(labelЦентр_secondClick);
        }

        private void labelЦентр_secondClick(object sender, EventArgs e)
        {
            labelЛево.BackColor = Color.DarkSeaGreen;
            richТекст.SelectionAlignment = HorizontalAlignment.Left;
            labelЛево.Click += new EventHandler(labelЛево_firstClick);
            labelЛево.Click -= new EventHandler(labelЛево_secondClick);
        }

        private void labelЛево_firstClick(object sender, EventArgs e)
        {
            labelЛево.BackColor = Color.SeaGreen;
            richТекст.SelectionAlignment = HorizontalAlignment.Left;
            labelЛево.Click -= new EventHandler(labelЛево_firstClick);
            labelЛево.Click += new EventHandler(labelЛево_secondClick);
        }

        private void labelЛево_secondClick(object sender, EventArgs e)
        {
            labelЛево.BackColor = Color.DarkSeaGreen;
            richТекст.SelectionAlignment = HorizontalAlignment.Left;
            labelЛево.Click += new EventHandler(labelЛево_firstClick);
            labelЛево.Click -= new EventHandler(labelЛево_secondClick);
        }

        private void labelЗаголовок_firstClick(object sender, EventArgs e)
        {
            labelЗаголовок.BackColor = Color.Khaki;
            SendKeys.Send("{ENTER}");
            richТекст.SelectionAlignment = HorizontalAlignment.Center;
            richТекст.SelectionColor = Color.DodgerBlue;
            FontStyle style = (FontStyle.Bold);//жирный
            richТекст.SelectionFont = new Font("Tahoma", 16, style);
            labelЗаголовок.Click -= new EventHandler(labelЗаголовок_firstClick); //отключаем первый обработчик
            labelЗаголовок.Click += new EventHandler(labelЗаголовок_secondClick);//включаем второй обработчик
        }

        private void labelЗаголовок_secondClick(object sender, EventArgs e)
        {
            labelЗаголовок.BackColor = Color.Khaki;
            SendKeys.Send("{ENTER}");
            richТекст.SelectionAlignment = HorizontalAlignment.Center;
            richТекст.SelectionColor = Color.Black;
            FontStyle style = (FontStyle.Regular);
            richТекст.SelectionFont = new Font("Microsoft Sans Serif", 10, style);
            labelЗаголовок.Click += new EventHandler(labelЗаголовок_firstClick); //включаем первый обработчик
            labelЗаголовок.Click -= new EventHandler(labelЗаголовок_secondClick);//отключаем второй обработчик
        }

        private void labelЦвет_firstClick(object sender, EventArgs e)
        {
            panelЦвета.Visible = true;
            labelЦвет.Click -= new EventHandler(labelЦвет_firstClick);
            labelЦвет.Click += new EventHandler(labelЦвет_secondClick);
        }

        private void labelЦвет_secondClick(object sender, EventArgs e)
        {
            panelЦвета.Visible = false;
            labelЦвет.Click += new EventHandler(labelЦвет_firstClick);
            labelЦвет.Click -= new EventHandler(labelЦвет_secondClick);
        }

        private void labelКурсив_firstClick(object sender, EventArgs e)
        {
            labelКурсив.BackColor = Color.SeaGreen;
            int newFontSize = 10; //размер
            FontStyle style = (FontStyle.Italic); //курсив
            richТекст.SelectionFont = new Font(richТекст.Font.FontFamily, (float)newFontSize, style);
            labelКурсив.Click -= new EventHandler(labelКурсив_firstClick);
            labelКурсив.Click += new EventHandler(labelКурсив_secondClick);
        }

        private void labelКурсив_secondClick(object sender, EventArgs e)
        {
            labelКурсив.BackColor = Color.DarkSeaGreen;
            richТекст.SelectionFont = new Font("Microsoft Sans Serif", 10);
            labelКурсив.Click += new EventHandler(labelКурсив_firstClick);
            labelКурсив.Click -= new EventHandler(labelКурсив_secondClick);
        }

        private void labelЗачеркнутый_firstClick(object sender, EventArgs e)
        {
            labelЗачеркнутый.BackColor = Color.SeaGreen;
            int newFontSize = 10; //размер
            FontStyle style = (FontStyle.Strikeout); //зачеркнутый
            richТекст.SelectionFont = new Font(richТекст.Font.FontFamily, (float)newFontSize, style);
            labelЗачеркнутый.Click -= new EventHandler(labelЗачеркнутый_firstClick);
            labelЗачеркнутый.Click += new EventHandler(labelЗачеркнутый_secondClick);
        }

        private void labelЗачеркнутый_secondClick(object sender, EventArgs e)
        {
            labelЗачеркнутый.BackColor = Color.DarkSeaGreen;
            richТекст.SelectionFont = new Font("Microsoft Sans Serif", 10);
            labelЗачеркнутый.Click += new EventHandler(labelЗачеркнутый_firstClick);
            labelЗачеркнутый.Click -= new EventHandler(labelЗачеркнутый_secondClick);
        }

        private void labelПодчеркнутый_firstClick(object sender, EventArgs e)
        {
            labelПодчеркнутый.BackColor = Color.SeaGreen;
            int newFontSize = 10; //размер
            FontStyle style = (FontStyle.Underline); //подчеркнутый
            richТекст.SelectionFont = new Font(richТекст.Font.FontFamily, (float)newFontSize, style);
            labelПодчеркнутый.Click -= new EventHandler(labelПодчеркнутый_firstClick);
            labelПодчеркнутый.Click += new EventHandler(labelПодчеркнутый_secondClick);
        }

        private void labelПодчеркнутый_secondClick(object sender, EventArgs e)
        {
            labelПодчеркнутый.BackColor = Color.DarkSeaGreen;
            richТекст.SelectionFont = new Font("Microsoft Sans Serif", 10);
            labelПодчеркнутый.Click += new EventHandler(labelПодчеркнутый_firstClick);
            labelПодчеркнутый.Click -= new EventHandler(labelПодчеркнутый_secondClick);
        }

        private void labelПолужирный_firstClick(object sender, EventArgs e)
        {
            labelПолужирный.BackColor = Color.SeaGreen;
            int newFontSize = 10; //размер
            FontStyle style = (FontStyle.Bold); //жирный
            richТекст.SelectionFont = new Font(richТекст.Font.FontFamily, (float)newFontSize, style);
            labelПолужирный.Click -= new EventHandler(labelПолужирный_firstClick);
            labelПолужирный.Click += new EventHandler(labelПолужирный_secondClick);
        }

        private void labelПолужирный_secondClick(object sender, EventArgs e)
        {
            labelПолужирный.BackColor = Color.DarkSeaGreen;
            richТекст.SelectionFont = new Font("Microsoft Sans Serif", 10);
            labelПолужирный.Click += new EventHandler(labelПолужирный_firstClick);
            labelПолужирный.Click -= new EventHandler(labelПолужирный_secondClick);
        }

        private void labelРазмер_firstClick(object sender, EventArgs e)
        {
            panelРазмер.Visible = true;
            labelРазмер.Click -= new EventHandler(labelРазмер_firstClick);
            labelРазмер.Click += new EventHandler(labelРазмер_secondClick);
        }

        private void labelРазмер_secondClick(object sender, EventArgs e)
        {
            panelРазмер.Visible = false;
            labelРазмер.Click += new EventHandler(labelРазмер_firstClick);
            labelРазмер.Click -= new EventHandler(labelРазмер_secondClick);
        }

        private void labelЖирный_firstClick(object sender, EventArgs e)
        {
            labelЖирный.BackColor = Color.SeaGreen;
            richТекст.SelectionFont = new Font("Arial Black", 11F,
                System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));//задаем стиль шрифта
            labelЖирный.Click -= new EventHandler(labelЖирный_firstClick);
            labelЖирный.Click += new EventHandler(labelЖирный_secondClick);
        }

        private void labelЖирный_secondClick(object sender, EventArgs e)
        {
            labelЖирный.BackColor = Color.DarkSeaGreen;
            richТекст.SelectionFont = new Font("Microsoft Sans Serif", 10);
            labelЖирный.Click += new EventHandler(labelЖирный_firstClick);
            labelЖирный.Click -= new EventHandler(labelЖирный_secondClick);
        }

        private void btnДобавить1_Click(object sender, EventArgs e)
        {
            cueДатаОтправки.Enabled = true;
            cueОтправитель1.Enabled = true;
            cueПочтовоеОтделение.Enabled = true;
            cueПолучатель1.Enabled = true;
            cueВидПересылки.Enabled = true;
            cueТрекНомер1.Enabled = true;
            cueСумма1.Enabled = true;
            comboВыборПересылки.Enabled = true;
            comboТрекНомер1.Enabled = true;
            //добавляем новую строку в таблицу "Отправка" базы данных "МояПосылка"
            try
            {
                cueДатаОтправки.Focus();
                this.мояПосылкаDataSet.Отправка.AddОтправкаRow(this.мояПосылкаDataSet.Отправка.NewОтправкаRow());
                отправкаBindingSource.MoveLast();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Сообщение", MessageBoxButtons.OK, MessageBoxIcon.Error);
                отправкаBindingSource.ResetBindings(false);
            }
        }

        private void btnСохранить1_Click(object sender, EventArgs e)
        {
            if (cueДатаОтправки.Text == string.Empty)
            {
                MessageBox.Show("Заполните пожалуйста поле Дата отправки", "Информация", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }
            if (cueОтправитель1.Text == string.Empty)
            {
                MessageBox.Show("Заполните пожалуйста поле Отправитель", "Информация", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }
            if (cueПочтовоеОтделение.Text == string.Empty)
            {
                MessageBox.Show("Заполните пожалуйста поле Почтовое отделение", "Информация", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }
            if (cueПолучатель1.Text == string.Empty)
            {
                MessageBox.Show("Заполните пожалуйста поле Получатель", "Информация", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }
            if (cueВидОтправления.Text == string.Empty)
            {
                MessageBox.Show("Выберите пожалуйста Вид почтового отправления", "Информация", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }
            if (cueВидПересылки.Text == string.Empty)
            {
                MessageBox.Show("Заполните пожалуйста поле Вид пересылки", "Информация", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }
            if (cueТрекНомер1.Text == string.Empty)
            {
                MessageBox.Show("Заполните пожалуйста поле Номер почтового идентификатора", "Информация", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }
            if (cueСумма1.Text == string.Empty)
            {
                MessageBox.Show("Заполните пожалуйста поле Сумма оплаты за пересылку", "Информация", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }
            //сохраняем добавленную строку в таблице "Отправка" базы данных "МояПосылка"
            try
            {
                отправкаBindingSource.EndEdit();
                отправкаTableAdapter.Update(this.мояПосылкаDataSet.Отправка);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Сообщение", MessageBoxButtons.OK, MessageBoxIcon.Error);
                отправкаBindingSource.ResetBindings(false);
            }
            //суммируем значения ячеек последнего столбца advancedDataGridView. Полученное значение выводим в текстовое поле, как строковое выражение
            try
            {
                double balans = 0;
                foreach (DataGridViewRow row in advancedDataGridView.Rows)
                {
                    double incom;
                    double.TryParse((row.Cells[8].Value ?? "0").ToString().Replace(".", ","), out incom);
                    balans += incom;
                }
                txtСумма.Text = balans.ToString();
            }
            catch (Exception)
            {
                MessageBox.Show("Заполните пожалуйста поле Сумма оплаты за пересылку!", "Ошибка в расчетах", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            groupРедактор1.Visible = false;
        }

        private void btnРедактировать1_Click(object sender, EventArgs e)
        {
            groupРедактор1.Visible = true;
            cueДатаОтправки.Enabled = true;
            cueОтправитель1.Enabled = true;
            cueПочтовоеОтделение.Enabled = true;
            cueПолучатель1.Enabled = true;
            cueВидПересылки.Enabled = true;
            cueТрекНомер1.Enabled = true;
            cueСумма1.Enabled = true;
            comboВыборПересылки.Enabled = true;
            comboТрекНомер1.Enabled = true;
        }

        private void dateTimePicker_ValueChanged(object sender, EventArgs e)
        {
            date = dateTimePicker.Value;
            cueДатаОтправки.Text = date.ToString("dd.MM.yyyy", CultureInfo.GetCultureInfo("ru-ru"));
        }

        private void comboВыборПересылки_SelectedIndexChanged(object sender, EventArgs e)
        {
            cueВидПересылки.Text = comboВыборПересылки.Text;
        }

        private void comboТрекНомер1_SelectedIndexChanged(object sender, EventArgs e)
        {
            cueТрекНомер1.Text = comboТрекНомер1.Text;
        }

        private void cueДатаОтправки_Enter(object sender, EventArgs e)
        {
            cueДатаОтправки.BackColor = Color.Honeydew;
            dateTimePicker.Visible = true;
        }

        private void cueДатаОтправки_Leave(object sender, EventArgs e)
        {
            cueДатаОтправки.BackColor = Color.PaleGreen;
        }

        private void cueОтправитель1_Enter(object sender, EventArgs e)
        {
            cueОтправитель1.BackColor = Color.Honeydew;
            dateTimePicker.Visible = false;
        }

        private void cueОтправитель1_Leave(object sender, EventArgs e)
        {
            cueОтправитель1.BackColor = Color.PaleGreen;
        }

        private void cueПочтовоеОтделение_Enter(object sender, EventArgs e)
        {
            cueПочтовоеОтделение.BackColor = Color.Honeydew;
            dateTimePicker.Visible = false;
        }

        private void cueПочтовоеОтделение_Leave(object sender, EventArgs e)
        {
            cueПочтовоеОтделение.BackColor = Color.PaleGreen;
        }

        private void cueПолучатель1_Enter(object sender, EventArgs e)
        {
            cueПолучатель1.BackColor = Color.Honeydew;
            dateTimePicker.Visible = false;
        }

        private void cueПолучатель1_Leave(object sender, EventArgs e)
        {
            cueПолучатель1.BackColor = Color.PaleGreen;
        }

        private void cueВидПересылки_Enter(object sender, EventArgs e)
        {
            cueВидПересылки.BackColor = Color.Honeydew;
            dateTimePicker.Visible = false;
        }

        private void cueВидПересылки_Leave(object sender, EventArgs e)
        {
            cueВидПересылки.BackColor = Color.PaleGreen;
        }

        private void cueТрекНомер1_Enter(object sender, EventArgs e)
        {
            cueТрекНомер1.BackColor = Color.Honeydew;
            dateTimePicker.Visible = false;
        }

        private void cueТрекНомер1_Leave(object sender, EventArgs e)
        {
            cueТрекНомер1.BackColor = Color.PaleGreen;
        }

        private void cueСумма1_Enter(object sender, EventArgs e)
        {
            cueСумма1.BackColor = Color.Honeydew;
            dateTimePicker.Visible = false;
        }

        private void cueСумма1_Leave(object sender, EventArgs e)
        {
            cueСумма1.BackColor = Color.PaleGreen;
        }

        private void cueДатаОтправки_KeyUp(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
                cueОтправитель1.Focus();
        }

        private void cueОтправитель1_KeyUp(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
                cueПочтовоеОтделение.Focus();
        }

        private void cueПочтовоеОтделение_KeyUp(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
                cueПолучатель1.Focus();
        }

        private void cueПолучатель1_KeyUp(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
                cueВидПересылки.Focus();
        }

        private void cueВидПересылки_KeyUp(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
                cueТрекНомер1.Focus();
        }

        private void cueТрекНомер1_KeyUp(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
                cueСумма1.Focus();
        }

        private void comboBox1_Enter(object sender, EventArgs e)
        {
            comboBox1.BackColor = Color.Honeydew;
        }

        private void comboBox1_Leave(object sender, EventArgs e)
        {
            comboBox1.BackColor = Color.DarkSeaGreen;
        }

        private void comboBox2_Enter(object sender, EventArgs e)
        {
            comboBox2.BackColor = Color.Honeydew;
        }

        private void comboBox2_Leave(object sender, EventArgs e)
        {
            comboBox2.BackColor = Color.DarkSeaGreen;
        }

        private void comboBox3_Enter(object sender, EventArgs e)
        {
            comboBox3.BackColor = Color.Honeydew;
        }

        private void comboBox3_Leave(object sender, EventArgs e)
        {
            comboBox3.BackColor = Color.DarkSeaGreen;
        }

        private void comboBox4_Enter(object sender, EventArgs e)
        {
            comboBox4.BackColor = Color.Honeydew;
        }

        private void comboBox4_Leave(object sender, EventArgs e)
        {
            comboBox4.BackColor = Color.DarkSeaGreen;
        }

        private void btnМеждународное_Click(object sender, EventArgs e)
        {
            cueВидОтправления.Text = "Международное";
        }

        private void btnПоРоссии_Click(object sender, EventArgs e)
        {
            cueВидОтправления.Text = "По России";
        }

        private void advancedDataGridView_FilterStringChanged(object sender, EventArgs e)
        {
            this.отправкаBindingSource.Filter = this.advancedDataGridView.FilterString;
        }

        private void advancedDataGridView_SortStringChanged(object sender, EventArgs e)
        {
            this.отправкаBindingSource.Filter = this.advancedDataGridView.SortString;
        }

        private void advancedDataGridView_KeyDown(object sender, KeyEventArgs e)
        {
            //удаляем все значения из таблицы "Отправка"
            if (e.KeyCode == Keys.Delete)
                if (MessageBox.Show("Вы уверены, что хотите очистить все данные?", "Очистка данных", MessageBoxButtons.OKCancel, MessageBoxIcon.Information) == DialogResult.OK)
                {
                    отправкаBindingSource.RemoveCurrent();
                }
        }

        private void cueПоиск1_Enter(object sender, EventArgs e)
        {
            cueПоиск1.BackColor = Color.Honeydew;
        }

        private void cueПоиск1_Leave(object sender, EventArgs e)
        {
            cueПоиск1.BackColor = Color.PaleGreen;
        }

        private void btnПоиск1_Click(object sender, EventArgs e)
        {
            for (int i = 0; i < advancedDataGridView.RowCount; i++)
            {
                advancedDataGridView.Rows[i].Selected = false;
                for (int j = 0; j < advancedDataGridView.ColumnCount; j++)
                    if (advancedDataGridView.Rows[i].Cells[j].Value != null)
                        if (advancedDataGridView.Rows[i].Cells[j].Value.ToString().Contains(cueПоиск1.Text))
                        {
                            advancedDataGridView.Rows[i].Selected = true;
                            break;
                        }
            }
            cueПоиск1.Text = "";
        }

        private void btnОтменаПоиска2_Click(object sender, EventArgs e)
        {
            advancedDataGridView.ClearSelection();
            cueПоиск1.Text = "";
        }

        private void linkДобавить_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            cueДатаОтправки.Enabled = true;
            cueОтправитель1.Enabled = true;
            cueПочтовоеОтделение.Enabled = true;
            cueПолучатель1.Enabled = true;
            cueВидПересылки.Enabled = true;
            cueТрекНомер1.Enabled = true;
            cueСумма1.Enabled = true;
            comboВыборПересылки.Enabled = true;
            comboТрекНомер1.Enabled = true;
            try
            {
                cueДатаОтправки.Focus();
                this.мояПосылкаDataSet.Отправка.AddОтправкаRow(this.мояПосылкаDataSet.Отправка.NewОтправкаRow());
                отправкаBindingSource.MoveLast();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Сообщение", MessageBoxButtons.OK, MessageBoxIcon.Error);
                отправкаBindingSource.ResetBindings(false);
            }
        }

        private void linkСохранить_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            if (cueДатаОтправки.Text == string.Empty)
            {
                MessageBox.Show("Заполните пожалуйста поле Дата отправки", "Информация", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }
            if (cueОтправитель1.Text == string.Empty)
            {
                MessageBox.Show("Заполните пожалуйста поле Отправитель", "Информация", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }
            if (cueПочтовоеОтделение.Text == string.Empty)
            {
                MessageBox.Show("Заполните пожалуйста поле Почтовое отделение", "Информация", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }
            if (cueПолучатель1.Text == string.Empty)
            {
                MessageBox.Show("Заполните пожалуйста поле Получатель", "Информация", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }
            if (cueВидОтправления.Text == string.Empty)
            {
                MessageBox.Show("Выберите пожалуйста Вид почтового отправления", "Информация", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }
            if (cueВидПересылки.Text == string.Empty)
            {
                MessageBox.Show("Заполните пожалуйста поле Вид пересылки", "Информация", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }
            if (cueТрекНомер1.Text == string.Empty)
            {
                MessageBox.Show("Заполните пожалуйста поле Номер почтового идентификатора", "Информация", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }
            if (cueСумма1.Text == string.Empty)
            {
                MessageBox.Show("Заполните пожалуйста поле Сумма оплаты за пересылку", "Информация", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }
            try
            {
                отправкаBindingSource.EndEdit();
                отправкаTableAdapter.Update(this.мояПосылкаDataSet.Отправка);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Сообщение", MessageBoxButtons.OK, MessageBoxIcon.Error);
                отправкаBindingSource.ResetBindings(false);
            }
            try
            {
                double balans = 0;
                foreach (DataGridViewRow row in advancedDataGridView.Rows)
                {
                    double incom;
                    double.TryParse((row.Cells[8].Value ?? "0").ToString().Replace(".", ","), out incom);
                    balans += incom;
                }
                txtСумма.Text = balans.ToString();
            }
            catch (Exception)
            {
                MessageBox.Show("Заполните пожалуйста поле Сумма оплаты за пересылку!", "Ошибка в расчетах", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            groupРедактор1.Visible = false;
        }

        private void linkРедактировать_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            groupРедактор1.Visible = true;
            cueДатаОтправки.Enabled = true;
            cueОтправитель1.Enabled = true;
            cueПочтовоеОтделение.Enabled = true;
            cueПолучатель1.Enabled = true;
            cueВидПересылки.Enabled = true;
            cueТрекНомер1.Enabled = true;
            cueСумма1.Enabled = true;
            comboВыборПересылки.Enabled = true;
            comboТрекНомер1.Enabled = true;
        }

        private void linkПоиск_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            groupПоиск1.Visible = true;
        }

        private void linkУдалить_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            //чтобы не выскакивало исключение "Нарушение параллелизма в таблице "Отправка" базы данных "МояПосылка" заполняем удаляемую строку данными и сохраняем ее
            if (MessageBox.Show("Вы уверены, что хотите удалить данную посылку?", "Удаление данных", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
            {
                cueДатаОтправки.Text = "01.01.1900";
                cueОтправитель1.Text = "Посылка удалена";
                cueПолучатель1.Text = "Посылка удалена";
                cueПочтовоеОтделение.Text = "Посылка удалена";
                cueВидОтправления.Text = "Посылка удалена";
                cueВидПересылки.Text = "Посылка удалена";
                cueТрекНомер1.Text = "Посылка удалена";
                cueСумма1.Text = "00,0";
                try
                {
                    отправкаBindingSource.EndEdit();
                    отправкаTableAdapter.Update(this.мояПосылкаDataSet.Отправка);
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message, "Сообщение", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    отправкаBindingSource.ResetBindings(false);
                }
                try
                {
                    double balans = 0;
                    foreach (DataGridViewRow row in advancedDataGridView.Rows)
                    {
                        double incom;
                        double.TryParse((row.Cells[8].Value ?? "0").ToString().Replace(".", ","), out incom);
                        balans += incom;
                    }
                    txtСумма.Text = balans.ToString();
                }
                catch (Exception)
                {
                    MessageBox.Show("Заполните пожалуйста поле Сумма оплаты за пересылку!", "Ошибка в расчетах", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
                //окрашиваем строку по условию в розовый цвет
                for (int i = 0; i < advancedDataGridView.RowCount; i++)
                {
                    for (int j = 0; j < advancedDataGridView.Columns.Count; j++)
                    {

                        switch (advancedDataGridView[j, i].FormattedValue.ToString())
                        {
                            case "Посылка удалена":
                                advancedDataGridView[j, i].Style.BackColor = Color.LightPink;
                                advancedDataGridView[j, i].Style.ForeColor = Color.LightPink;
                                break;
                        }
                    }
                }
            }

        }

        private void linkЗакончитьВвод_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            cueДатаОтправки.Enabled = false;
            cueОтправитель1.Enabled = false;
            cueПочтовоеОтделение.Enabled = false;
            cueПолучатель1.Enabled = false;
            cueВидПересылки.Enabled = false;
            cueТрекНомер1.Enabled = false;
            cueСумма1.Enabled = false;
            comboВыборПересылки.Enabled = false;
            comboТрекНомер1.Enabled = false;
        }

        Excel.Application exApp_New = new Excel.Application();
        Excel.Workbook wb_New = null;
        Excel.Worksheet ws_New = null;

        private void linkЕксель_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            saveFileDialog1.Filter = "Excel files (*.xls)|*.xls|Excel files (*.xlsx)|*.xlsx";
            saveFileDialog1.FilterIndex = 2;
            saveFileDialog1.FileName = "Почтовые отправления";
            saveFileDialog1.Title = "Сохранение документа";
            saveFileDialog1.RestoreDirectory = true;

            if (saveFileDialog1.ShowDialog() == DialogResult.OK)
            {
                try
                {
                    wb_New = exApp_New.Workbooks.Add(System.Reflection.Missing.Value);
                    ws_New = (Microsoft.Office.Interop.Excel.Worksheet)wb_New.Worksheets.get_Item(1);
                    ws_New.Cells.Locked = false;
                    Microsoft.Office.Interop.Excel.Range rangeWidth1 = ws_New.Range["A1", System.Type.Missing];
                    rangeWidth1.EntireColumn.ColumnWidth = 5;
                    Excel.Range rangeWidth2 = ws_New.Range["B1", System.Type.Missing];
                    rangeWidth2.EntireColumn.ColumnWidth = 17;
                    Excel.Range rangeWidth3 = ws_New.Range["C1", System.Type.Missing];
                    rangeWidth3.EntireColumn.ColumnWidth = 44;
                    Excel.Range rangeWidth4 = ws_New.Range["D1", System.Type.Missing];
                    rangeWidth4.EntireColumn.ColumnWidth = 44;
                    Excel.Range rangeWidth5 = ws_New.Range["E1", System.Type.Missing];
                    rangeWidth5.EntireColumn.ColumnWidth = 44;
                    Excel.Range rangeWidth6 = ws_New.Range["F1", System.Type.Missing];
                    rangeWidth6.EntireColumn.ColumnWidth = 27;
                    Excel.Range rangeWidth7 = ws_New.Range["G1", System.Type.Missing];
                    rangeWidth7.EntireColumn.ColumnWidth = 17;
                    Excel.Range rangeWidth8 = ws_New.Range["H1", System.Type.Missing];
                    rangeWidth8.EntireColumn.ColumnWidth = 33;
                    Excel.Range rangeWidth9 = ws_New.Range["I1", System.Type.Missing];
                    rangeWidth9.EntireColumn.ColumnWidth = 27;
                    ws_New.Cells[1, 1] = "№";
                    ws_New.Cells[1, 2] = "Дата отправки";
                    ws_New.Cells[1, 3] = "Отправитель";
                    ws_New.Cells[1, 4] = "Почтовое отделение";
                    ws_New.Cells[1, 5] = "Получатель";
                    ws_New.Cells[1, 6] = "Вид почтового отправления";
                    ws_New.Cells[1, 7] = "Вид пересылки";
                    ws_New.Cells[1, 8] = "Номер почтового идентификатора";
                    ws_New.Cells[1, 9] = "Сумма оплаты за пересылку";

                    for (int i = 0; i < advancedDataGridView.ColumnCount; i++)
                    {
                        for (int j = 0; j < advancedDataGridView.RowCount; j++)
                        {
                            ws_New.Cells[j + 2, i + 1] = (advancedDataGridView[i, j].Value).ToString();
                        }
                    }
                    Excel.Range tRange = ws_New.UsedRange;
                    tRange.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                    tRange.Borders.Weight = Excel.XlBorderWeight.xlThin;
                    Excel.Range cellRange = (Excel.Range)ws_New.Cells[1, 1];
                    Excel.Range rowRange = cellRange.EntireRow;
                    rowRange.Insert(Excel.XlInsertShiftDirection.xlShiftDown, false);
                    Microsoft.Office.Interop.Excel.Range Табель = (Microsoft.Office.Interop.Excel.Range)ws_New.Cells[1, 1];
                    Табель.Value2 = txtШапка.Text;
                    exApp_New.Visible = true;
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message, "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }

        private void linkТХТ_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            Stream myStream;

            saveFileDialog1.Filter = "Текстовый файл (*.txt)|*.txt";
            saveFileDialog1.FilterIndex = 1;
            saveFileDialog1.FileName = "Почтовые отправления";
            saveFileDialog1.Title = "Сохранение документа";
            saveFileDialog1.RestoreDirectory = true;

            if (saveFileDialog1.ShowDialog() == DialogResult.OK)
            {
                if ((myStream = saveFileDialog1.OpenFile()) != null)
                {
                    StreamWriter myWritet = new StreamWriter(myStream);
                    myWritet.WriteLine(txtШапка.Text);
                    try
                    {
                        for (int i = 0; i < advancedDataGridView.RowCount; i++)
                        {
                            for (int j = 0; j < advancedDataGridView.ColumnCount; j++)
                            {
                                myWritet.Write(advancedDataGridView.Rows[i].Cells[j].Value.ToString() + " ");
                            }
                            myWritet.WriteLine();
                        }
                        
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(ex.Message);
                    }
                    finally
                    {
                        myWritet.Close();
                    }
                }
            }
            System.Diagnostics.Process.Start(saveFileDialog1.FileName);
        }

        private void linkОчиститьВсе_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            if (MessageBox.Show("Если вы хотите очистить все данные, выделите в таблице все строки и нажмите на клавиатуре клавишу DELETE!", "Сообщение", MessageBoxButtons.OK, MessageBoxIcon.Information) == DialogResult.OK)
            {
                advancedDataGridView.Visible = true;
            }
        }

        private void linkПоказТаблицы_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            advancedDataGridView.Visible = true;
        }

        private void linkСоздатьШапку_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            panelШапка.Visible = true;
            panelШапка.Size = new Size(493, 188);
            panelШапка.BackColor = Color.DarkSeaGreen;
            panelШапка.BorderStyle = BorderStyle.FixedSingle;
            panelШапка.Location = new Point(227, 7);
        }

        private void FormГлавная_FormClosing(object sender, FormClosingEventArgs e)
        {
            Properties.Settings.Default.txtШапка = txtШапка.Text;
            Properties.Settings.Default.txtШапка2 = txtШапка2.Text;
            Properties.Settings.Default.richТекст = richТекст.Text;
            Properties.Settings.Default.Save();
        }

        private void txtШапка_Enter(object sender, EventArgs e)
        {
            txtШапка.BackColor = Color.Honeydew;
        }

        private void txtШапка_Leave(object sender, EventArgs e)
        {
            txtШапка.BackColor = Color.PaleGreen;
        }

        private void btnЗакрыть_Click(object sender, EventArgs e)
        {
            panelШапка.Size = new Size(248, 19);
            panelШапка.BackColor = Color.DarkSeaGreen;
            panelШапка.BorderStyle = BorderStyle.None;
            panelШапка.Location = new Point(700, 6);
            panelШапка.Visible = false;
        }

        private void btnДобавить2_Click(object sender, EventArgs e)
        {
            cueДатаПолучения.Enabled = true;
            cueПолучатель2.Enabled = true;
            cueПочтовоеОтделение2.Enabled = true;
            cueОтправитель2.Enabled = true;
            cueТрекНомер2.Enabled = true;
            comboТрекНомер2.Enabled = true;
            try
            {
                cueДатаПолучения.Focus();
                this.мояПосылка2DataSet.Получение.AddПолучениеRow(this.мояПосылка2DataSet.Получение.NewПолучениеRow());
                получениеBindingSource.MoveLast();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Сообщение", MessageBoxButtons.OK, MessageBoxIcon.Error);
                получениеBindingSource.ResetBindings(false);
            }
        }

        private void btnМеждународное2_Click(object sender, EventArgs e)
        {
            txtВидОтправления.Text = "Международное";
        }

        private void btnПоРоссии2_Click(object sender, EventArgs e)
        {
            txtВидОтправления.Text = "По России";
        }

        private void btnСохранить2_Click(object sender, EventArgs e)
        {
            if (cueДатаПолучения.Text == string.Empty)
            {
                MessageBox.Show("Заполните пожалуйста поле Дата получения", "Информация", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }
            if (cueПолучатель2.Text == string.Empty)
            {
                MessageBox.Show("Заполните пожалуйста поле Получатель", "Информация", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }
            if (cueПочтовоеОтделение2.Text == string.Empty)
            {
                MessageBox.Show("Заполните пожалуйста поле Почтовое отделение", "Информация", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }
            if (cueОтправитель2.Text == string.Empty)
            {
                MessageBox.Show("Заполните пожалуйста поле Отправитель", "Информация", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }
            if (txtВидОтправления.Text == string.Empty)
            {
                MessageBox.Show("Выберите пожалуйста Вид почтового отправления", "Информация", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }
            if (cueТрекНомер2.Text == string.Empty)
            {
                MessageBox.Show("Заполните пожалуйста поле Номер почтового идентификатора", "Информация", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }
            try
            {
                получениеBindingSource.EndEdit();
                получениеTableAdapter.Update(this.мояПосылка2DataSet.Получение);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Сообщение", MessageBoxButtons.OK, MessageBoxIcon.Error);
                получениеBindingSource.ResetBindings(false);
            }
            groupРедактор2.Visible = false;
        }

        private void btnРедактировать2_Click(object sender, EventArgs e)
        {
            groupРедактор2.Visible = true;
            cueДатаПолучения.Enabled = true;
            cueПолучатель2.Enabled = true;
            cueПочтовоеОтделение2.Enabled = true;
            cueОтправитель2.Enabled = true;
            cueТрекНомер2.Enabled = true;
            comboТрекНомер2.Enabled = true;
        }

        private void btnНайти2_Click(object sender, EventArgs e)
        {
            for (int i = 0; i < advancedDataGridView2.RowCount; i++)
            {
                advancedDataGridView2.Rows[i].Selected = false;
                for (int j = 0; j < advancedDataGridView2.ColumnCount; j++)
                    if (advancedDataGridView2.Rows[i].Cells[j].Value != null)
                        if (advancedDataGridView2.Rows[i].Cells[j].Value.ToString().Contains(cueПоиск2.Text))
                        {
                            advancedDataGridView2.Rows[i].Selected = true;
                            break;
                        }
            }
            cueПоиск2.Text = "";
        }

        private void btnОтменитьПоиск2_Click(object sender, EventArgs e)
        {
            advancedDataGridView2.ClearSelection();
            cueПоиск2.Text = "";
        }

        private void advancedDataGridView2_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Delete)
                if (MessageBox.Show("Вы уверены, что хотите очистить все данные?", "Очистка данных", MessageBoxButtons.OKCancel, MessageBoxIcon.Information) == DialogResult.OK)
                {
                    получениеBindingSource.RemoveCurrent();
                }
        }

        private void advancedDataGridView2_FilterStringChanged(object sender, EventArgs e)
        {
            this.получениеBindingSource.Filter = this.advancedDataGridView.FilterString;
        }

        private void advancedDataGridView2_SortStringChanged(object sender, EventArgs e)
        {
            this.получениеBindingSource.Filter = this.advancedDataGridView.SortString;
        }

        private void cueДатаПолучения_Enter(object sender, EventArgs e)
        {
            cueДатаПолучения.BackColor = Color.Honeydew;
            dateTimePicker1.Visible = true;
        }

        private void cueДатаПолучения_Leave(object sender, EventArgs e)
        {
            cueДатаПолучения.BackColor = Color.PaleGreen;
        }

        private void cueПолучатель2_Enter(object sender, EventArgs e)
        {
            cueПолучатель2.BackColor = Color.Honeydew;
            dateTimePicker1.Visible = false;
        }

        private void cueПолучатель2_Leave(object sender, EventArgs e)
        {
            cueПолучатель2.BackColor = Color.PaleGreen;
        }

        private void cueПочтовоеОтделение2_Enter(object sender, EventArgs e)
        {
            cueПочтовоеОтделение2.BackColor = Color.Honeydew;
            dateTimePicker1.Visible = false;
        }

        private void cueПочтовоеОтделение2_Leave(object sender, EventArgs e)
        {
            cueПочтовоеОтделение2.BackColor = Color.PaleGreen;
        }

        private void cueОтправитель2_Enter(object sender, EventArgs e)
        {
            cueОтправитель2.BackColor = Color.Honeydew;
            dateTimePicker1.Visible = false;
        }

        private void cueОтправитель2_Leave(object sender, EventArgs e)
        {
            cueОтправитель2.BackColor = Color.PaleGreen;
        }

        private void cueТрекНомер2_Enter(object sender, EventArgs e)
        {
            cueТрекНомер2.BackColor = Color.Honeydew;
            dateTimePicker1.Visible = false;
        }

        private void cueТрекНомер2_Leave(object sender, EventArgs e)
        {
            cueТрекНомер2.BackColor = Color.PaleGreen;
        }

        private void cueДатаПолучения_KeyUp(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
                cueПолучатель2.Focus();
        }

        private void cueПолучатель2_KeyUp(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
                cueПочтовоеОтделение2.Focus();
        }

        private void cueПочтовоеОтделение2_KeyUp(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
                cueОтправитель2.Focus();
        }

        private void cueОтправитель2_KeyUp(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
                cueТрекНомер2.Focus();
        }

        private void cueПоиск2_Enter(object sender, EventArgs e)
        {
            cueПоиск2.BackColor = Color.Honeydew;
        }

        private void cueПоиск2_Leave(object sender, EventArgs e)
        {
            cueПоиск2.BackColor = Color.PaleGreen;
        }

        private void linkДобавить2_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            cueДатаПолучения.Enabled = true;
            cueПолучатель2.Enabled = true;
            cueПочтовоеОтделение2.Enabled = true;
            cueОтправитель2.Enabled = true;
            cueТрекНомер2.Enabled = true;
            comboТрекНомер2.Enabled = true;
            try
            {
                cueДатаПолучения.Focus();
                this.мояПосылка2DataSet.Получение.AddПолучениеRow(this.мояПосылка2DataSet.Получение.NewПолучениеRow());
                получениеBindingSource.MoveLast();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Сообщение", MessageBoxButtons.OK, MessageBoxIcon.Error);
                получениеBindingSource.ResetBindings(false);
            }
        }

        private void linkСохранить2_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            if (cueДатаПолучения.Text == string.Empty)
            {
                MessageBox.Show("Заполните пожалуйста поле Дата получения", "Информация", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }
            if (cueПолучатель2.Text == string.Empty)
            {
                MessageBox.Show("Заполните пожалуйста поле Получатель", "Информация", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }
            if (cueПочтовоеОтделение2.Text == string.Empty)
            {
                MessageBox.Show("Заполните пожалуйста поле Почтовое отделение", "Информация", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }
            if (cueОтправитель2.Text == string.Empty)
            {
                MessageBox.Show("Заполните пожалуйста поле Отправитель", "Информация", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }
            if (txtВидОтправления.Text == string.Empty)
            {
                MessageBox.Show("Выберите пожалуйста Вид почтового отправления", "Информация", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }
            if (cueТрекНомер2.Text == string.Empty)
            {
                MessageBox.Show("Заполните пожалуйста поле Номер почтового идентификатора", "Информация", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }
            try
            {
                получениеBindingSource.EndEdit();
                получениеTableAdapter.Update(this.мояПосылка2DataSet.Получение);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Сообщение", MessageBoxButtons.OK, MessageBoxIcon.Error);
                получениеBindingSource.ResetBindings(false);
            }
            groupРедактор2.Visible = false;
        }

        private void linkРедактировать2_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            groupРедактор2.Visible = true;
            cueДатаПолучения.Enabled = true;
            cueПолучатель2.Enabled = true;
            cueПочтовоеОтделение2.Enabled = true;
            cueОтправитель2.Enabled = true;
            cueТрекНомер2.Enabled = true;
            comboТрекНомер2.Enabled = true;
        }

        private void linkПоиск2_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            groupПоиск2.Visible = true;
        }

        private void linkУдалить2_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            if (MessageBox.Show("Вы уверены, что хотите удалить данную посылку?", "Удаление данных", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
            {
                cueДатаПолучения.Text = "01.01.1900";
                cueПолучатель2.Text = "Посылка удалена";
                cueПочтовоеОтделение2.Text = "Посылка удалена";
                cueОтправитель2.Text = "Посылка удалена";
                txtВидОтправления.Text = "Посылка удалена";
                cueТрекНомер2.Text = "Посылка удалена";
                try
                {
                    получениеBindingSource.EndEdit();
                    получениеTableAdapter.Update(this.мояПосылка2DataSet.Получение);
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message, "Сообщение", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    получениеBindingSource.ResetBindings(false);
                }
                for (int i = 0; i < advancedDataGridView2.RowCount; i++)
                {
                    for (int j = 0; j < advancedDataGridView2.Columns.Count; j++)
                    {

                        switch (advancedDataGridView2[j, i].FormattedValue.ToString())
                        {
                            case "Посылка удалена":
                                advancedDataGridView2[j, i].Style.BackColor = Color.LightPink;
                                advancedDataGridView2[j, i].Style.ForeColor = Color.LightPink;
                                break;
                        }
                    }
                }
            }
        }
    

        private void linkЗакончитьВвод2_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            cueДатаПолучения.Enabled = false;
            cueПолучатель2.Enabled = false;
            cueПочтовоеОтделение2.Enabled = false;
            cueОтправитель2.Enabled = false;
            cueТрекНомер2.Enabled = false;
            comboТрекНомер2.Enabled = false;
        }

        Excel.Application exApp2_New = new Excel.Application();
        Excel.Workbook wb2_New = null;
        Excel.Worksheet ws2_New = null;

        private void linkЕксель2_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            saveFileDialog1.Filter = "Excel files (*.xls)|*.xls|Excel files (*.xlsx)|*.xlsx";
            saveFileDialog1.FilterIndex = 2;
            saveFileDialog1.FileName = "Получение почты";
            saveFileDialog1.Title = "Сохранение документа";
            saveFileDialog1.RestoreDirectory = true;

            if (saveFileDialog1.ShowDialog() == DialogResult.OK)
            {
                try
                {
                    wb2_New = exApp2_New.Workbooks.Add(System.Reflection.Missing.Value);
                    ws2_New = (Microsoft.Office.Interop.Excel.Worksheet)wb2_New.Worksheets.get_Item(1);
                    ws2_New.Cells.Locked = false;
                    Microsoft.Office.Interop.Excel.Range rangeWidth1 = ws2_New.Range["A1", System.Type.Missing];
                    rangeWidth1.EntireColumn.ColumnWidth = 5;
                    Excel.Range rangeWidth2 = ws2_New.Range["B1", System.Type.Missing];
                    rangeWidth2.EntireColumn.ColumnWidth = 17;
                    Excel.Range rangeWidth3 = ws2_New.Range["C1", System.Type.Missing];
                    rangeWidth3.EntireColumn.ColumnWidth = 44;
                    Excel.Range rangeWidth4 = ws2_New.Range["D1", System.Type.Missing];
                    rangeWidth4.EntireColumn.ColumnWidth = 44;
                    Excel.Range rangeWidth5 = ws2_New.Range["E1", System.Type.Missing];
                    rangeWidth5.EntireColumn.ColumnWidth = 44;
                    Excel.Range rangeWidth6 = ws2_New.Range["F1", System.Type.Missing];
                    rangeWidth6.EntireColumn.ColumnWidth = 27;
                    Excel.Range rangeWidth7 = ws2_New.Range["G1", System.Type.Missing];
                    rangeWidth7.EntireColumn.ColumnWidth = 30;
                    ws2_New.Cells[1, 1] = "№";
                    ws2_New.Cells[1, 2] = "Дата получения";
                    ws2_New.Cells[1, 3] = "Получатель";
                    ws2_New.Cells[1, 4] = "Почтовое отделение получателя";
                    ws2_New.Cells[1, 5] = "Отправитель";
                    ws2_New.Cells[1, 6] = "Вид почтового отправления";
                    ws2_New.Cells[1, 7] = "Номер почтового идентификатора";

                    for (int i = 0; i < advancedDataGridView2.ColumnCount; i++)
                    {
                        for (int j = 0; j < advancedDataGridView2.RowCount; j++)
                        {
                            ws2_New.Cells[j + 2, i + 1] = (advancedDataGridView2[i, j].Value).ToString();
                        }
                    }
                    Excel.Range tRange = ws_New.UsedRange;
                    tRange.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                    tRange.Borders.Weight = Excel.XlBorderWeight.xlThin;
                    Excel.Range cellRange = (Excel.Range)ws_New.Cells[1, 1];
                    Excel.Range rowRange = cellRange.EntireRow;
                    rowRange.Insert(Excel.XlInsertShiftDirection.xlShiftDown, false);
                    Microsoft.Office.Interop.Excel.Range Табель = (Microsoft.Office.Interop.Excel.Range)ws_New.Cells[1, 1];
                    Табель.Value2 = txtШапка2.Text;
                    exApp2_New.Visible = true;
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message, "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }

        private void linkТХТ2_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            Stream myStream;

            saveFileDialog1.Filter = "Текстовый файл (*.txt)|*.txt";
            saveFileDialog1.FilterIndex = 1;
            saveFileDialog1.FileName = "Получение почты";
            saveFileDialog1.Title = "Сохранение документа";
            saveFileDialog1.RestoreDirectory = true;

            if (saveFileDialog1.ShowDialog() == DialogResult.OK)
            {
                if ((myStream = saveFileDialog1.OpenFile()) != null)
                {
                    StreamWriter myWritet = new StreamWriter(myStream);
                    myWritet.WriteLine(txtШапка2.Text);
                    try
                    {
                        for (int i = 0; i < advancedDataGridView2.RowCount; i++)
                        {
                            for (int j = 0; j < advancedDataGridView2.ColumnCount; j++)
                            {
                                myWritet.Write(advancedDataGridView2.Rows[i].Cells[j].Value.ToString() + " ");
                            }
                            myWritet.WriteLine();
                        }

                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(ex.Message);
                    }
                    finally
                    {
                        myWritet.Close();
                    }
                }
            }
            System.Diagnostics.Process.Start(saveFileDialog1.FileName);
        }

        private void linkОчиститьВсе2_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            if (MessageBox.Show("Если вы хотите очистить все данные, выделите в таблице все строки и нажмите на клавиатуре клавишу DELETE!", "Сообщение", MessageBoxButtons.OK, MessageBoxIcon.Information) == DialogResult.OK)
            {
                advancedDataGridView2.Visible = true;
            }
        }

        private void linkПоказатьТаблицу2_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            advancedDataGridView2.Visible = true;
        }

        private void linkШапка2_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            panelШапка2.Visible = true;
            panelШапка2.Size = new Size(472, 136);
            panelШапка2.BackColor = Color.DarkSeaGreen;
            panelШапка2.BorderStyle = BorderStyle.FixedSingle;
            panelШапка2.Location = new Point(227, 7);
        }

        private void btnЗакрыть2_Click(object sender, EventArgs e)
        {
            panelШапка2.Size = new Size(248, 19);
            panelШапка2.BackColor = Color.DarkSeaGreen;
            panelШапка2.BorderStyle = BorderStyle.None;
            panelШапка2.Location = new Point(700, 6);
            panelШапка2.Visible = false;
        }

        private void txtШапка2_Enter(object sender, EventArgs e)
        {
            txtШапка2.BackColor = Color.Honeydew;
        }

        private void txtШапка2_Leave(object sender, EventArgs e)
        {
            txtШапка2.BackColor = Color.PaleGreen;
        }

        private void dateTimePicker1_ValueChanged(object sender, EventArgs e)
        {
            date = dateTimePicker.Value;
            cueДатаПолучения.Text = date.ToString("dd.MM.yyyy", CultureInfo.GetCultureInfo("ru-ru"));
        }

        private void comboТрекНомер2_SelectedIndexChanged(object sender, EventArgs e)
        {
            cueТрекНомер2.Text = comboТрекНомер2.Text;
        }

        private void comboBox8_Enter(object sender, EventArgs e)
        {
            comboBox8.BackColor = Color.Honeydew;
        }

        private void comboBox8_Leave(object sender, EventArgs e)
        {
            comboBox8.BackColor = Color.DarkSeaGreen;
        }

        private void comboBox7_Enter(object sender, EventArgs e)
        {
            comboBox7.BackColor = Color.Honeydew;
        }

        private void comboBox7_Leave(object sender, EventArgs e)
        {
            comboBox7.BackColor = Color.DarkSeaGreen;
        }

        private void comboBox6_Enter(object sender, EventArgs e)
        {
            comboBox6.BackColor = Color.Honeydew;
        }

        private void comboBox6_Leave(object sender, EventArgs e)
        {
            comboBox6.BackColor = Color.DarkSeaGreen;
        }

        private void comboBox5_Enter(object sender, EventArgs e)
        {
            comboBox5.BackColor = Color.Honeydew;
        }

        private void comboBox5_Leave(object sender, EventArgs e)
        {
            comboBox5.BackColor = Color.DarkSeaGreen;
        }

        private void treeWeb_AfterSelect(object sender, TreeViewEventArgs e)
        {
            switch (e.Node.Name)
            {
                case "Узел0":
                    {
                        webBrowser.Navigate("https://www.pochta.ru/parcels");
                    }
                    break;
                case "Узел1":
                    {
                        webBrowser.Navigate("https://www.pochta.ru/parcels");
                    }
                    break;
                case "Узел2":
                    {
                        webBrowser.Navigate("https://www.pochta.ru/form?type=F7P");
                    }
                    break;
                case "Узел3":
                    {
                        webBrowser.Navigate("https://www.pochta.ru/form?type=BANDEROL_ADDRESS_LABEL");
                    }
                    break;
                case "Узел4":
                    {
                        webBrowser.Navigate("https://www.pochta.ru/form?type=CN22");
                    }
                    break;
                case "Узел5":
                    {
                        webBrowser.Navigate("https://www.pochta.ru/form?type=F7A");
                    }
                    break;
                case "Узел6":
                    {
                        webBrowser.Navigate("https://www.pochta.ru/form?type=PETIT_ADDRESS_LABEL");
                    }
                    break;
                case "Узел7":
                    {
                        webBrowser.Navigate("https://www.pochta.ru/form?type=CN22");
                    }
                    break;
                case "Узел8":
                    {
                        webBrowser.Navigate("https://www.pochta.ru/form?type=F7P");
                    }
                    break;
                case "Узел9":
                    {
                        webBrowser.Navigate("https://www.pochta.ru/form?type=F116");
                    }
                    break;
                case "Узел10":
                    {
                        webBrowser.Navigate("https://www.pochta.ru/form?type=CP71");
                    }
                    break;
                case "Узел11":
                    {
                        webBrowser.Navigate("https://www.pochta.ru/form?type=CN23");
                    }
                    break;
                case "Узел12":
                    {
                        webBrowser.Navigate("https://www.pochta.ru/form?type=F107");
                    }
                    break;
                case "Узел13":
                    {
                        webBrowser.Navigate("https://www.pochta.ru/form?type=F22");
                    }
                    break;
                case "Узел14":
                    {
                        webBrowser.Navigate("https://www.pochta.ru/tracking");
                    }
                    break;
                case "Узел15":
                    {
                        webBrowser.Navigate("https://www.pochta.ru/form?type=CLAIM_INTERNAL");
                    }
                    break;
                case "Узел16":
                    {
                        webBrowser.Navigate("https://www.pochta.ru/form?type=CLAIM_INTERNATIONAL");
                    }
                    break;
                case "Узел17":
                    {
                        webBrowser.Navigate("https://www.pochta.ru/claim");
                    }
                    break;
            }
        }

        private void labelПрикрепить_MouseMove(object sender, MouseEventArgs e)
        {
            labelПрикрепить.BackColor = Color.SeaGreen;
        }

        private void labelПрикрепить_MouseLeave(object sender, EventArgs e)
        {
            labelПрикрепить.BackColor = Color.DarkSeaGreen;
        }

        private void labelОтменить_MouseMove(object sender, MouseEventArgs e)
        {
            labelОтменить.BackColor = Color.SeaGreen;
        }

        private void labelОтменить_MouseLeave(object sender, EventArgs e)
        {
            labelОтменить.BackColor = Color.DarkSeaGreen;
        }

        private void labelВернуть_MouseMove(object sender, MouseEventArgs e)
        {
            labelВернуть.BackColor = Color.SeaGreen;
        }

        private void labelВернуть_MouseLeave(object sender, EventArgs e)
        {
            labelВернуть.BackColor = Color.DarkSeaGreen;
        }

        private void labelРазмер_MouseMove(object sender, MouseEventArgs e)
        {
            labelРазмер.BackColor = Color.SeaGreen;
        }

        private void labelРазмер_MouseLeave(object sender, EventArgs e)
        {
            labelРазмер.BackColor = Color.DarkSeaGreen;
        }

        private void labelЖирный_MouseMove(object sender, MouseEventArgs e)
        {
            labelЖирный.BackColor = Color.SeaGreen;
        }

        private void labelЖирный_MouseLeave(object sender, EventArgs e)
        {
            labelЖирный.BackColor = Color.DarkSeaGreen;
        }

        private void labelПолужирный_MouseMove(object sender, MouseEventArgs e)
        {
            labelПолужирный.BackColor = Color.SeaGreen;
        }

        private void labelПолужирный_MouseLeave(object sender, EventArgs e)
        {
            labelПолужирный.BackColor = Color.DarkSeaGreen;
        }

        private void labelПодчеркнутый_MouseMove(object sender, MouseEventArgs e)
        {
            labelПодчеркнутый.BackColor = Color.SeaGreen;
        }

        private void labelПодчеркнутый_MouseLeave(object sender, EventArgs e)
        {
            labelПодчеркнутый.BackColor = Color.DarkSeaGreen;
        }

        private void labelЗачеркнутый_MouseMove(object sender, MouseEventArgs e)
        {
            labelЗачеркнутый.BackColor = Color.SeaGreen;
        }

        private void labelЗачеркнутый_MouseLeave(object sender, EventArgs e)
        {
            labelЗачеркнутый.BackColor = Color.DarkSeaGreen;
        }

        private void labelКурсив_MouseMove(object sender, MouseEventArgs e)
        {
            labelКурсив.BackColor = Color.SeaGreen;
        }

        private void labelКурсив_MouseLeave(object sender, EventArgs e)
        {
            labelКурсив.BackColor = Color.DarkSeaGreen;
        }

        private void labelЦвет_MouseMove(object sender, MouseEventArgs e)
        {
            labelЦвет.BackColor = Color.SeaGreen;
        }

        private void labelЦвет_MouseLeave(object sender, EventArgs e)
        {
            labelЦвет.BackColor = Color.DarkSeaGreen;
        }

        private void labelЗаголовок_MouseMove(object sender, MouseEventArgs e)
        {
            labelЗаголовок.BackColor = Color.SeaGreen;
        }

        private void labelЗаголовок_MouseLeave(object sender, EventArgs e)
        {
            labelЗаголовок.BackColor = Color.DarkSeaGreen;
        }

        private void labelЛево_MouseMove(object sender, MouseEventArgs e)
        {
            labelЛево.BackColor = Color.SeaGreen;
        }

        private void labelЛево_MouseLeave(object sender, EventArgs e)
        {
            labelЛево.BackColor = Color.DarkSeaGreen;
        }

        private void labelЦентр_MouseMove(object sender, MouseEventArgs e)
        {
            labelЦентр.BackColor = Color.SeaGreen;
        }

        private void labelЦентр_MouseLeave(object sender, EventArgs e)
        {
            labelЦентр.BackColor = Color.DarkSeaGreen;
        }

        private void labelПраво_MouseMove(object sender, MouseEventArgs e)
        {
            labelПраво.BackColor = Color.SeaGreen;
        }

        private void labelПраво_MouseLeave(object sender, EventArgs e)
        {
            labelПраво.BackColor = Color.DarkSeaGreen;
        }

        private void labelВыделитьВсе_MouseMove(object sender, MouseEventArgs e)
        {
            labelВыделитьВсе.BackColor = Color.SeaGreen;
        }

        private void labelВыделитьВсе_MouseLeave(object sender, EventArgs e)
        {
            labelВыделитьВсе.BackColor = Color.DarkSeaGreen;
        }

        private void labelВырезать_MouseMove(object sender, MouseEventArgs e)
        {
            labelВырезать.BackColor = Color.SeaGreen;
        }

        private void labelВырезать_MouseLeave(object sender, EventArgs e)
        {
            labelВырезать.BackColor = Color.DarkSeaGreen;
        }

        private void labelГиперссылка_MouseMove(object sender, MouseEventArgs e)
        {
            labelГиперссылка.BackColor = Color.SeaGreen;
        }

        private void labelГиперссылка_MouseLeave(object sender, EventArgs e)
        {
            labelГиперссылка.BackColor = Color.DarkSeaGreen;
        }

        private void labelСохранитьКак_MouseMove(object sender, MouseEventArgs e)
        {
            labelСохранитьКак.BackColor = Color.SeaGreen;
        }

        private void labelСохранитьКак_MouseLeave(object sender, EventArgs e)
        {
            labelСохранитьКак.BackColor = Color.DarkSeaGreen;
        }

        private void labelУдалить_MouseMove(object sender, MouseEventArgs e)
        {
            labelУдалить.BackColor = Color.SeaGreen;
        }

        private void labelУдалить_MouseLeave(object sender, EventArgs e)
        {
            labelУдалить.BackColor = Color.DarkSeaGreen;
        }

        private void pictureФото_MouseMove(object sender, MouseEventArgs e)
        {
            lblФото.Visible = true;
        }

        private void pictureФото_MouseLeave(object sender, EventArgs e)
        {
            lblФото.Visible = false;
        }

        private void pictureФото_Click(object sender, EventArgs e)
        {
            try
            {
                using (OpenFileDialog ofd = new OpenFileDialog() { Filter = "Файлы JPG(*.JPG*)|*.JPG| Файлы PNG(*.png*)|*.png|Файлы BMP(*.bmp*)|*.bmp*", ValidateNames = true, Multiselect = false })
                {
                    if (ofd.ShowDialog() == DialogResult.OK)
                        pictureФото.Image = Image.FromFile(ofd.FileName);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void LblАвторизоваться_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            comboЛогин.Enabled = true;
            cueЛогин.Enabled = true;
            cueПароль.Enabled = true;
            label80.Visible = true;
            pictureФото.Visible = true;
            try
            {
                cueЛогин.Focus();
                this.мояПосылка3DataSet.ЛогинПароль.AddЛогинПарольRow(this.мояПосылка3DataSet.ЛогинПароль.NewЛогинПарольRow());
                логинПарольBindingSource.MoveLast();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Сообщение", MessageBoxButtons.OK, MessageBoxIcon.Error);
                логинПарольBindingSource.ResetBindings(false);
            }
        }

        private void lblGmail_Click(object sender, EventArgs e)
        {
            if (cueЛогин.Text == string.Empty)
            {
                MessageBox.Show("Заполните пожалуйста поле Логин!", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return;
            }
            if (cueПароль.Text == string.Empty)
            {
                MessageBox.Show("Заполните пожалуйста поле Пароль!", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return;
            }
            try
            {
                логинПарольBindingSource.EndEdit();
                логинПарольTableAdapter.Update(this.мояПосылка3DataSet.ЛогинПароль);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Сообщение", MessageBoxButtons.OK, MessageBoxIcon.Error);
                логинПарольBindingSource.ResetBindings(false);
            }
            panelАвторизация.Visible = false;
            treeEmail.Visible = true;
            panelОтправка.Visible = true;
            btnGmail.Visible = true;
        }

        private void lblOutlook_Click(object sender, EventArgs e)
        {
            if (cueЛогин.Text == string.Empty)
            {
                MessageBox.Show("Заполните пожалуйста поле Логин!", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return;
            }
            if (cueПароль.Text == string.Empty)
            {
                MessageBox.Show("Заполните пожалуйста поле Пароль!", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return;
            }
            try
            {
                логинПарольBindingSource.EndEdit();
                логинПарольTableAdapter.Update(this.мояПосылка3DataSet.ЛогинПароль);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Сообщение", MessageBoxButtons.OK, MessageBoxIcon.Error);
                логинПарольBindingSource.ResetBindings(false);
            }
            panelАвторизация.Visible = false;
            treeEmail.Visible = true;
            panelОтправка.Visible = true;
            btnOutlook.Visible = true;
        }

        private void lblYahoo_Click(object sender, EventArgs e)
        {
            if (cueЛогин.Text == string.Empty)
            {
                MessageBox.Show("Заполните пожалуйста поле Логин!", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return;
            }
            if (cueПароль.Text == string.Empty)
            {
                MessageBox.Show("Заполните пожалуйста поле Пароль!", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return;
            }
            try
            {
                логинПарольBindingSource.EndEdit();
                логинПарольTableAdapter.Update(this.мояПосылка3DataSet.ЛогинПароль);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Сообщение", MessageBoxButtons.OK, MessageBoxIcon.Error);
                логинПарольBindingSource.ResetBindings(false);
            }
            panelАвторизация.Visible = false;
            treeEmail.Visible = true;
            panelОтправка.Visible = true;
            btnYahoo.Visible = true;
        }

        private void lblMail_Click(object sender, EventArgs e)
        {
            if (cueЛогин.Text == string.Empty)
            {
                MessageBox.Show("Заполните пожалуйста поле Логин!", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return;
            }
            if (cueПароль.Text == string.Empty)
            {
                MessageBox.Show("Заполните пожалуйста поле Пароль!", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return;
            }
            try
            {
                логинПарольBindingSource.EndEdit();
                логинПарольTableAdapter.Update(this.мояПосылка3DataSet.ЛогинПароль);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Сообщение", MessageBoxButtons.OK, MessageBoxIcon.Error);
                логинПарольBindingSource.ResetBindings(false);
            }
            panelАвторизация.Visible = false;
            treeEmail.Visible = true;
            panelОтправка.Visible = true;
            btnMail.Visible = true;
        }

        private void lblYandex_Click(object sender, EventArgs e)
        {
            if (cueЛогин.Text == string.Empty)
            {
                MessageBox.Show("Заполните пожалуйста поле Логин!", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return;
            }
            if (cueПароль.Text == string.Empty)
            {
                MessageBox.Show("Заполните пожалуйста поле Пароль!", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return;
            }
            try
            {
                логинПарольBindingSource.EndEdit();
                логинПарольTableAdapter.Update(this.мояПосылка3DataSet.ЛогинПароль);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Сообщение", MessageBoxButtons.OK, MessageBoxIcon.Error);
                логинПарольBindingSource.ResetBindings(false);
            }
            panelАвторизация.Visible = false;
            treeEmail.Visible = true;
            panelОтправка.Visible = true;
            btnYandex.Visible = true;
        }

        private void lblRambler_Click(object sender, EventArgs e)
        {
            if (cueЛогин.Text == string.Empty)
            {
                MessageBox.Show("Заполните пожалуйста поле Логин!", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return;
            }
            if (cueПароль.Text == string.Empty)
            {
                MessageBox.Show("Заполните пожалуйста поле Пароль!", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return;
            }
            try
            {
                логинПарольBindingSource.EndEdit();
                логинПарольTableAdapter.Update(this.мояПосылка3DataSet.ЛогинПароль);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Сообщение", MessageBoxButtons.OK, MessageBoxIcon.Error);
                логинПарольBindingSource.ResetBindings(false);
            }
            panelАвторизация.Visible = false;
            treeEmail.Visible = true;
            panelОтправка.Visible = true;
            btnRambler.Visible = true;
        }

        private void comboЛогин_Enter(object sender, EventArgs e)
        {
            comboЛогин.BackColor = Color.Honeydew;
        }

        private void comboЛогин_Leave(object sender, EventArgs e)
        {
            comboЛогин.BackColor = Color.DarkSeaGreen;
        }

        private void cueЛогин_Enter(object sender, EventArgs e)
        {
            cueЛогин.BackColor = Color.Honeydew;
        }

        private void cueЛогин_Leave(object sender, EventArgs e)
        {
            cueЛогин.BackColor = Color.PaleGreen;
        }

        private void cueПароль_Enter(object sender, EventArgs e)
        {
            cueПароль.BackColor = Color.Honeydew;
        }

        private void cueПароль_Leave(object sender, EventArgs e)
        {
            cueПароль.BackColor = Color.PaleGreen;
        }

        private void btnНаписать_Click(object sender, EventArgs e)
        {
            отправкаEMailBindingSource.Filter = null;
            panelОтправкаЕмайл.Visible = true;
            panelИнструменты.Visible = true;
            panelСохранить.Visible = true;
            richТекст.Visible = true;
            try
            {
                richТекст.Focus();
                this.мояПосылка4DataSet.ОтправкаEMail.AddОтправкаEMailRow(this.мояПосылка4DataSet.ОтправкаEMail.NewОтправкаEMailRow());
                отправкаEMailBindingSource.MoveLast();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Сообщение", MessageBoxButtons.OK, MessageBoxIcon.Error);
                отправкаEMailBindingSource.ResetBindings(false);
            }
            txtДатаЕмайл.Text = DateTime.Now.ToString("dd.MM.yyyy");
            cueОтКого.Text = cueЛогин.Text;
            pictureФотоОтправителя.Image = pictureФото.Image;
            comboЗаполнить.Visible = false;
            comboЗаполнить2.Visible = false;
        }

        private void timer_Tick(object sender, EventArgs e)
        {
            labelСообщение.Visible = false;
        }

        private void btnGmail_Click(object sender, EventArgs e)
        {
            txtСтатус.Text = "Отправленные";
            if (cueОтКого.Text == string.Empty)
            {
                MessageBox.Show("Заполните пожалуйста поле От кого!", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return;
            }
            if (cueКому.Text == string.Empty)
            {
                MessageBox.Show("Заполните пожалуйста поле Кому!", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return;
            }
            try
            {
                отправкаEMailBindingSource.EndEdit();
                отправкаEMailTableAdapter.Update(this.мояПосылка4DataSet.ОтправкаEMail);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Сообщение", MessageBoxButtons.OK, MessageBoxIcon.Error);
                отправкаEMailBindingSource.ResetBindings(false);
            }
            try
            {
                SmtpClient client = new SmtpClient("smtp.gmail.com", 587);
                MailMessage message = new MailMessage();
                message.From = new MailAddress(cueОтКого.Text);
                message.To.Add(cueКому.Text);
                message.Body = richТекст.Text;
                message.Subject = cueТемаПисьма.Text;
                client.UseDefaultCredentials = false;
                client.DeliveryMethod = SmtpDeliveryMethod.Network;
                client.EnableSsl = true;
                if (txtПрикрепитьФайл.Text != "")
                {
                    message.Attachments.Add(new Attachment(txtПрикрепитьФайл.Text));
                }
                client.Credentials = new System.Net.NetworkCredential(cueОтКого.Text, cueПароль.Text);
                client.Send(message);
                message = null;
            }
            catch (Exception)
            {
                MessageBox.Show("При отправке сообщения произошла ошибка!", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            labelСообщение.Visible = true;
            labelСообщение.Text = "Посылка отправлена";
            comboЗаполнить.Visible = false;
            comboЗаполнить2.Visible = false;
        }

        private void btnOutlook_Click(object sender, EventArgs e)
        {
            txtСтатус.Text = "Отправленные";
            if (cueОтКого.Text == string.Empty)
            {
                MessageBox.Show("Заполните пожалуйста поле От кого!", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return;
            }
            if (cueКому.Text == string.Empty)
            {
                MessageBox.Show("Заполните пожалуйста поле Кому!", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return;
            }
            try
            {
                отправкаEMailBindingSource.EndEdit();
                отправкаEMailTableAdapter.Update(this.мояПосылка4DataSet.ОтправкаEMail);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Сообщение", MessageBoxButtons.OK, MessageBoxIcon.Error);
                отправкаEMailBindingSource.ResetBindings(false);
            }
            try
            {
                SmtpClient client = new SmtpClient("smtp.live.com", 587);
                MailMessage message = new MailMessage();
                message.From = new MailAddress(cueОтКого.Text);
                message.To.Add(cueКому.Text);
                message.Body = richТекст.Text;
                message.Subject = cueТемаПисьма.Text;
                client.UseDefaultCredentials = false;
                client.DeliveryMethod = SmtpDeliveryMethod.Network;
                client.EnableSsl = true;
                if (txtПрикрепитьФайл.Text != "")
                {
                    message.Attachments.Add(new Attachment(txtПрикрепитьФайл.Text));
                }
                client.Credentials = new System.Net.NetworkCredential(cueОтКого.Text, cueПароль.Text);
                client.Send(message);
                message = null;
            }
            catch (Exception)
            {
                MessageBox.Show("При отправке сообщения произошла ошибка!", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            labelСообщение.Visible = true;
            labelСообщение.Text = "Посылка отправлена";
            comboЗаполнить.Visible = false;
            comboЗаполнить2.Visible = false;
        }

        private void btnYahoo_Click(object sender, EventArgs e)
        {
            txtСтатус.Text = "Отправленные";
            if (cueОтКого.Text == string.Empty)
            {
                MessageBox.Show("Заполните пожалуйста поле От кого!", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return;
            }
            if (cueКому.Text == string.Empty)
            {
                MessageBox.Show("Заполните пожалуйста поле Кому!", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return;
            }
            try
            {
                отправкаEMailBindingSource.EndEdit();
                отправкаEMailTableAdapter.Update(this.мояПосылка4DataSet.ОтправкаEMail);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Сообщение", MessageBoxButtons.OK, MessageBoxIcon.Error);
                отправкаEMailBindingSource.ResetBindings(false);
            }
            try
            {
                SmtpClient client = new SmtpClient("smtp.mail.yahoo.com", 465);
                MailMessage message = new MailMessage();
                message.From = new MailAddress(cueОтКого.Text);
                message.To.Add(cueКому.Text);
                message.Body = richТекст.Text;
                message.Subject = cueТемаПисьма.Text;
                client.UseDefaultCredentials = false;
                client.DeliveryMethod = SmtpDeliveryMethod.Network;
                client.EnableSsl = true;
                if (txtПрикрепитьФайл.Text != "")
                {
                    message.Attachments.Add(new Attachment(txtПрикрепитьФайл.Text));
                }
                client.Credentials = new System.Net.NetworkCredential(cueОтКого.Text, cueПароль.Text);
                client.Send(message);
                message = null;
            }
            catch (Exception)
            {
                MessageBox.Show("При отправке сообщения произошла ошибка!", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            labelСообщение.Visible = true;
            labelСообщение.Text = "Посылка отправлена";
            comboЗаполнить.Visible = false;
            comboЗаполнить2.Visible = false;
        }

        private void btnMail_Click(object sender, EventArgs e)
        {
            txtСтатус.Text = "Отправленные";
            if (cueОтКого.Text == string.Empty)
            {
                MessageBox.Show("Заполните пожалуйста поле От кого!", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return;
            }
            if (cueКому.Text == string.Empty)
            {
                MessageBox.Show("Заполните пожалуйста поле Кому!", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return;
            }
            try
            {
                отправкаEMailBindingSource.EndEdit();
                отправкаEMailTableAdapter.Update(this.мояПосылка4DataSet.ОтправкаEMail);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Сообщение", MessageBoxButtons.OK, MessageBoxIcon.Error);
                отправкаEMailBindingSource.ResetBindings(false);
            }
            try
            {
                SmtpClient client = new SmtpClient("smtp.mail.com", 25);
                MailMessage message = new MailMessage();
                message.From = new MailAddress(cueОтКого.Text);
                message.To.Add(cueКому.Text);
                message.Body = richТекст.Text;
                message.Subject = cueТемаПисьма.Text;
                client.UseDefaultCredentials = false;
                client.DeliveryMethod = SmtpDeliveryMethod.Network;
                client.EnableSsl = true;
                if (txtПрикрепитьФайл.Text != "")
                {
                    message.Attachments.Add(new Attachment(txtПрикрепитьФайл.Text));
                }
                client.Credentials = new System.Net.NetworkCredential(cueОтКого.Text, cueПароль.Text);
                client.Send(message);
                message = null;
            }
            catch (Exception)
            {
                MessageBox.Show("При отправке сообщения произошла ошибка!", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            labelСообщение.Visible = true;
            labelСообщение.Text = "Посылка отправлена";
            comboЗаполнить.Visible = false;
            comboЗаполнить2.Visible = false;
        }

        private void btnYandex_Click(object sender, EventArgs e)
        {
            txtСтатус.Text = "Отправленные";
            if (cueОтКого.Text == string.Empty)
            {
                MessageBox.Show("Заполните пожалуйста поле От кого!", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return;
            }
            if (cueКому.Text == string.Empty)
            {
                MessageBox.Show("Заполните пожалуйста поле Кому!", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return;
            }
            try
            {
                отправкаEMailBindingSource.EndEdit();
                отправкаEMailTableAdapter.Update(this.мояПосылка4DataSet.ОтправкаEMail);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Сообщение", MessageBoxButtons.OK, MessageBoxIcon.Error);
                отправкаEMailBindingSource.ResetBindings(false);
            }
            try
            {
                SmtpClient client = new SmtpClient("smtp.yandex.com", 465);
                MailMessage message = new MailMessage();
                message.From = new MailAddress(cueОтКого.Text);
                message.To.Add(cueКому.Text);
                message.Body = richТекст.Text;
                message.Subject = cueТемаПисьма.Text;
                client.UseDefaultCredentials = false;
                client.DeliveryMethod = SmtpDeliveryMethod.Network;
                client.EnableSsl = true;
                if (txtПрикрепитьФайл.Text != "")
                {
                    message.Attachments.Add(new Attachment(txtПрикрепитьФайл.Text));
                }
                client.Credentials = new System.Net.NetworkCredential(cueОтКого.Text, cueПароль.Text);
                client.Send(message);
                message = null;
            }
            catch (Exception)
            {
                MessageBox.Show("При отправке сообщения произошла ошибка!", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            labelСообщение.Visible = true;
            labelСообщение.Text = "Посылка отправлена";
            comboЗаполнить.Visible = false;
            comboЗаполнить2.Visible = false;
        }

        private void btnRambler_Click(object sender, EventArgs e)
        {
            txtСтатус.Text = "Отправленные";
            if (cueОтКого.Text == string.Empty)
            {
                MessageBox.Show("Заполните пожалуйста поле От кого!", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return;
            }
            if (cueКому.Text == string.Empty)
            {
                MessageBox.Show("Заполните пожалуйста поле Кому!", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return;
            }
            try
            {
                отправкаEMailBindingSource.EndEdit();
                отправкаEMailTableAdapter.Update(this.мояПосылка4DataSet.ОтправкаEMail);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Сообщение", MessageBoxButtons.OK, MessageBoxIcon.Error);
                отправкаEMailBindingSource.ResetBindings(false);
            }
            try
            {
                SmtpClient client = new SmtpClient("smtp.rambler.com", 465);
                MailMessage message = new MailMessage();
                message.From = new MailAddress(cueОтКого.Text);
                message.To.Add(cueКому.Text);
                message.Body = richТекст.Text;
                message.Subject = cueТемаПисьма.Text;
                client.UseDefaultCredentials = false;
                client.DeliveryMethod = SmtpDeliveryMethod.Network;
                client.EnableSsl = true;
                if (txtПрикрепитьФайл.Text != "")
                {
                    message.Attachments.Add(new Attachment(txtПрикрепитьФайл.Text));
                }
                client.Credentials = new System.Net.NetworkCredential(cueОтКого.Text, cueПароль.Text);
                client.Send(message);
                message = null;
            }
            catch (Exception)
            {
                MessageBox.Show("При отправке сообщения произошла ошибка!", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            labelСообщение.Visible = true;
            labelСообщение.Text = "Посылка отправлена";
            comboЗаполнить.Visible = false;
            comboЗаполнить2.Visible = false;
        }

        private void btnСохранитьПисьмо_Click(object sender, EventArgs e)
        {
            txtСтатус.Text = "Сохраненные";
            if (cueОтКого.Text == string.Empty)
            {
                MessageBox.Show("Заполните пожалуйста поле От кого!", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return;
            }
            if (cueКому.Text == string.Empty)
            {
                MessageBox.Show("Заполните пожалуйста поле Кому!", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return;
            }
            try
            {
                отправкаEMailBindingSource.EndEdit();
                отправкаEMailTableAdapter.Update(this.мояПосылка4DataSet.ОтправкаEMail);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Сообщение", MessageBoxButtons.OK, MessageBoxIcon.Error);
                отправкаEMailBindingSource.ResetBindings(false);
            }
            comboЗаполнить.Visible = false;
            comboЗаполнить2.Visible = false;
        }

        private void comboЛогин_SelectedIndexChanged(object sender, EventArgs e)
        {
            cueЛогин.Text = comboЛогин.Text;
        }

        private void cmbПолучатель_SelectedIndexChanged(object sender, EventArgs e)
        {
            cueКому.Text = cmbПолучатель.Text;
        }

        private void btnСоздатьШаблон_Click(object sender, EventArgs e)
        {
            отправкаEMailBindingSource.Filter = null;
            txtСтатус.Text = "Шаблоны";
            label7.Visible = true;
            try
            {
                richТекст.Focus();
                this.мояПосылка4DataSet.ОтправкаEMail.AddОтправкаEMailRow(this.мояПосылка4DataSet.ОтправкаEMail.NewОтправкаEMailRow());
                отправкаEMailBindingSource.MoveLast();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Сообщение", MessageBoxButtons.OK, MessageBoxIcon.Error);
                отправкаEMailBindingSource.ResetBindings(false);
            }
            txtДатаЕмайл.Text = DateTime.Now.ToString("dd.MM.yyyy");
            cueОтКого.Text = cueЛогин.Text;
            pictureBox1.Image = pictureФото.Image;
        }

        private void btnСохранитьШаблон_Click(object sender, EventArgs e)
        {
            txtСтатус.Text = "Шаблоны";
            if (cueОтКого.Text == string.Empty)
            {
                MessageBox.Show("Заполните пожалуйста поле От кого!", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return;
            }
            if (cueКому.Text == string.Empty)
            {
                MessageBox.Show("Заполните пожалуйста поле Кому!", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return;
            }
            try
            {
                отправкаEMailBindingSource.EndEdit();
                отправкаEMailTableAdapter.Update(this.мояПосылка4DataSet.ОтправкаEMail);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Сообщение", MessageBoxButtons.OK, MessageBoxIcon.Error);
                отправкаEMailBindingSource.ResetBindings(false);
            }
        }

        private void labelПрикрепить_Click(object sender, EventArgs e)
        {
            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                txtПрикрепитьФайл.Text = openFileDialog.FileName.ToString();
            }
            panelВложение.Visible = true;
        }

        private void cueОтКого_Enter(object sender, EventArgs e)
        {
            cueОтКого.BackColor = Color.Honeydew;
        }

        private void cueОтКого_Leave(object sender, EventArgs e)
        {
            cueОтКого.BackColor = Color.PaleGreen;
        }

        private void cueКому_Enter(object sender, EventArgs e)
        {
            cueКому.BackColor = Color.Honeydew;
        }

        private void cueКому_Leave(object sender, EventArgs e)
        {
            cueКому.BackColor = Color.PaleGreen;
        }

        private void cueТемаПисьма_Enter(object sender, EventArgs e)
        {
            cueТемаПисьма.BackColor = Color.Honeydew;
        }

        private void cueТемаПисьма_Leave(object sender, EventArgs e)
        {
            cueТемаПисьма.BackColor = Color.PaleGreen;
        }

        private void cueПоискПисем_Enter(object sender, EventArgs e)
        {
            cueПоискПисем.BackColor = Color.Honeydew;
        }

        private void cueПоискПисем_Leave(object sender, EventArgs e)
        {
            cueПоискПисем.BackColor = Color.PaleGreen;
        }

        private void cmbПолучатель_Enter(object sender, EventArgs e)
        {
            cmbПолучатель.BackColor = Color.Honeydew;
        }

        private void cmbПолучатель_Leave(object sender, EventArgs e)
        {
            cmbПолучатель.BackColor = Color.DarkSeaGreen;
        }

        private void richТекст_Enter(object sender, EventArgs e)
        {
            richТекст.BackColor = Color.Honeydew;
        }

        private void richТекст_Leave(object sender, EventArgs e)
        {
            richТекст.BackColor = Color.PaleGreen;
        }

        private void panelВложение_MouseEnter(object sender, EventArgs e)
        {
            labelОткрытьФайл.Visible = true;
        }

        private void panelВложение_MouseLeave(object sender, EventArgs e)
        {
            labelОткрытьФайл.Visible = false;
        }

        private void labelВложение_MouseEnter(object sender, EventArgs e)
        {
            labelОткрытьФайл.Visible = true;
        }

        private void labelОтменить_Click(object sender, EventArgs e)
        {
            richТекст.Undo();
        }

        private void labelВернуть_Click(object sender, EventArgs e)
        {
            richТекст.Redo();
        }

        private void labelВыделитьВсе_Click(object sender, EventArgs e)
        {
            richТекст.SelectAll();
        }

        private void richТекст_LinkClicked(object sender, LinkClickedEventArgs e)
        {
            System.Diagnostics.Process.Start(e.LinkText);
        }

        private void labelВырезать_Click(object sender, EventArgs e)
        {
            richТекст.SelectedText = "";
        }

        private void labelГиперссылка_Click(object sender, EventArgs e)
        {
            cueСсылка.Text = "";
            panelСсылка.Visible = true;
        }

        private void btnВставить_Click(object sender, EventArgs e)
        {
            richТекст.AppendText(cueСсылка.Text);
            panelСсылка.Visible = false;
        }

        private void btnОтмена_Click(object sender, EventArgs e)
        {
            cueСсылка.Text = "";
            panelСсылка.Visible = false;
        }

        private void labelУдалить_Click(object sender, EventArgs e)
        {
            txtСтатус.Text = "Удаленные";
            if (cueОтКого.Text == string.Empty)
            {
                MessageBox.Show("Заполните пожалуйста поле От кого!", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return;
            }
            if (cueКому.Text == string.Empty)
            {
                MessageBox.Show("Заполните пожалуйста поле Кому!", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return;
            }
            try
            {
                отправкаEMailBindingSource.EndEdit();
                отправкаEMailTableAdapter.Update(this.мояПосылка4DataSet.ОтправкаEMail);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Сообщение", MessageBoxButtons.OK, MessageBoxIcon.Error);
                отправкаEMailBindingSource.ResetBindings(false);
            }
        }

        private void lblВажные_Click(object sender, EventArgs e)
        {
            txtСтатус.Text = "Важные";
            if (cueОтКого.Text == string.Empty)
            {
                MessageBox.Show("Заполните пожалуйста поле От кого!", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return;
            }
            if (cueКому.Text == string.Empty)
            {
                MessageBox.Show("Заполните пожалуйста поле Кому!", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return;
            }
            try
            {
                отправкаEMailBindingSource.EndEdit();
                отправкаEMailTableAdapter.Update(this.мояПосылка4DataSet.ОтправкаEMail);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Сообщение", MessageBoxButtons.OK, MessageBoxIcon.Error);
                отправкаEMailBindingSource.ResetBindings(false);
            }
            panelСохранитьКак.Visible = false;
            comboЗаполнить.Visible = false;
            comboЗаполнить2.Visible = false;
        }

        private void lblСвложением_Click(object sender, EventArgs e)
        {
            txtСтатус.Text = "С вложением";
            if (cueОтКого.Text == string.Empty)
            {
                MessageBox.Show("Заполните пожалуйста поле От кого!", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return;
            }
            if (cueКому.Text == string.Empty)
            {
                MessageBox.Show("Заполните пожалуйста поле Кому!", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return;
            }
            try
            {
                отправкаEMailBindingSource.EndEdit();
                отправкаEMailTableAdapter.Update(this.мояПосылка4DataSet.ОтправкаEMail);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Сообщение", MessageBoxButtons.OK, MessageBoxIcon.Error);
                отправкаEMailBindingSource.ResetBindings(false);
            }
            panelСохранитьКак.Visible = false;
            comboЗаполнить.Visible = false;
            comboЗаполнить2.Visible = false;
        }

        private void lblЧерновик_Click(object sender, EventArgs e)
        {
            txtСтатус.Text = "Черновик";
            if (cueОтКого.Text == string.Empty)
            {
                MessageBox.Show("Заполните пожалуйста поле От кого!", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return;
            }
            if (cueКому.Text == string.Empty)
            {
                MessageBox.Show("Заполните пожалуйста поле Кому!", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return;
            }
            try
            {
                отправкаEMailBindingSource.EndEdit();
                отправкаEMailTableAdapter.Update(this.мояПосылка4DataSet.ОтправкаEMail);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Сообщение", MessageBoxButtons.OK, MessageBoxIcon.Error);
                отправкаEMailBindingSource.ResetBindings(false);
            }
            panelСохранитьКак.Visible = false;
            comboЗаполнить.Visible = false;
            comboЗаполнить2.Visible = false;
        }

        private void lblАрхив_Click(object sender, EventArgs e)
        {
            txtСтатус.Text = "Архив";
            if (cueОтКого.Text == string.Empty)
            {
                MessageBox.Show("Заполните пожалуйста поле От кого!", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return;
            }
            if (cueКому.Text == string.Empty)
            {
                MessageBox.Show("Заполните пожалуйста поле Кому!", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return;
            }
            try
            {
                отправкаEMailBindingSource.EndEdit();
                отправкаEMailTableAdapter.Update(this.мояПосылка4DataSet.ОтправкаEMail);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Сообщение", MessageBoxButtons.OK, MessageBoxIcon.Error);
                отправкаEMailBindingSource.ResetBindings(false);
            }
            panelСохранитьКак.Visible = false;
            comboЗаполнить.Visible = false;
            comboЗаполнить2.Visible = false;
        }

        private void label101_Click(object sender, EventArgs e)
        {
            richТекст.SelectionFont = new Font("Microsoft Sans Serif", 8);
            panelРазмер.Visible = false;
        }

        private void label100_Click(object sender, EventArgs e)
        {
            richТекст.SelectionFont = new Font("Microsoft Sans Serif", 9);
            panelРазмер.Visible = false;
        }

        private void label99_Click(object sender, EventArgs e)
        {
            richТекст.SelectionFont = new Font("Microsoft Sans Serif", 10);
            panelРазмер.Visible = false;
        }

        private void label98_Click(object sender, EventArgs e)
        {
            richТекст.SelectionFont = new Font("Microsoft Sans Serif", 12);
            panelРазмер.Visible = false;
        }

        private void label97_Click(object sender, EventArgs e)
        {
            richТекст.SelectionFont = new Font("Microsoft Sans Serif", 14);
            panelРазмер.Visible = false;
        }

        private void label96_Click(object sender, EventArgs e)
        {
            richТекст.SelectionFont = new Font("Microsoft Sans Serif", 16);
            panelРазмер.Visible = false;
        }

        private void label95_Click(object sender, EventArgs e)
        {
            richТекст.SelectionFont = new Font("Microsoft Sans Serif", 18);
            panelРазмер.Visible = false;
        }

        private void label94_Click(object sender, EventArgs e)
        {
            richТекст.SelectionFont = new Font("Microsoft Sans Serif", 20);
            panelРазмер.Visible = false;
        }

        private void label93_Click(object sender, EventArgs e)
        {
            richТекст.SelectionFont = new Font("Microsoft Sans Serif", 22);
            panelРазмер.Visible = false;
        }

        private void label92_Click(object sender, EventArgs e)
        {
            richТекст.SelectionFont = new Font("Microsoft Sans Serif", 24);
            panelРазмер.Visible = false;
        }

        private void label91_Click(object sender, EventArgs e)
        {
            richТекст.SelectionFont = new Font("Microsoft Sans Serif", 26);
            panelРазмер.Visible = false;
        }

        private void label90_Click(object sender, EventArgs e)
        {
            richТекст.SelectionFont = new Font("Microsoft Sans Serif", 28);
            panelРазмер.Visible = false;
        }

        private void labelОткрытьФайл_Click(object sender, EventArgs e)
        {
            
        }

        private void treeEmail_AfterSelect(object sender, TreeViewEventArgs e)
        {
            switch (e.Node.Name)
            {
                case "Узел0":
                    {
                        txtПоиск.Text = "Отправленные";
                        panelОтправкаЕмайл.Visible = true;
                        panelВложение.Visible = false;
                        richТекст.Visible = true;
                        отправкаEMailBindingSource.Filter = "Статус=\'" + txtПоиск.Text + "\'";
                        dataGridView.Visible = true;
                    }
                    break;
                case "Узел1":
                    {
                        txtПоиск.Text = "Сохраненные";
                        panelОтправкаЕмайл.Visible = true;
                        panelВложение.Visible = false;
                        richТекст.Visible = true;
                        отправкаEMailBindingSource.Filter = "Статус=\'" + txtПоиск.Text + "\'";
                        dataGridView.Visible = true;
                    }
                    break;
                case "Узел2":
                    {
                        txtПоиск.Text = "Важные";
                        panelОтправкаЕмайл.Visible = true;
                        panelВложение.Visible = false;
                        richТекст.Visible = true;
                        отправкаEMailBindingSource.Filter = "Статус=\'" + txtПоиск.Text + "\'";
                        dataGridView.Visible = true;
                    }
                    break;
                case "Узел3":
                    {
                        txtПоиск.Text = "С вложением";
                        panelОтправкаЕмайл.Visible = true;
                        panelВложение.Visible = true;
                        richТекст.Visible = true;
                        отправкаEMailBindingSource.Filter = "Статус=\'" + txtПоиск.Text + "\'";
                        dataGridView.Visible = true;
                    }
                    break;
                case "Узел4":
                    {
                        txtПоиск.Text = "Черновик";
                        panelОтправкаЕмайл.Visible = true;
                        panelВложение.Visible = false;
                        richТекст.Visible = true;
                        отправкаEMailBindingSource.Filter = "Статус=\'" + txtПоиск.Text + "\'";
                        dataGridView.Visible = true;
                    }
                    break;
                case "Узел5":
                    {
                        txtПоиск.Text = "Архив";
                        panelОтправкаЕмайл.Visible = true;
                        panelВложение.Visible = false;
                        richТекст.Visible = true;
                        отправкаEMailBindingSource.Filter = "Статус=\'" + txtПоиск.Text + "\'";
                        dataGridView.Visible = true;
                    }
                    break;
                case "Узел6":
                    {
                        txtПоиск.Text = "Шаблоны";
                        panelОтправкаЕмайл.Visible = true;
                        panelВложение.Visible = false;
                        richТекст.Visible = true;
                        отправкаEMailBindingSource.Filter = "Статус=\'" + txtПоиск.Text + "\'";
                        dataGridView.Visible = true;
                    }
                    break;
                case "Узел7":
                    {
                        txtПоиск.Text = "Удаленные";
                        panelОтправкаЕмайл.Visible = true;
                        panelВложение.Visible = false;
                        richТекст.Visible = true;
                        отправкаEMailBindingSource.Filter = "Статус=\'" + txtПоиск.Text + "\'";
                        dataGridView.Visible = true;
                    }
                    break;
            }
        }

        private void label78_Click(object sender, EventArgs e)
        {
            richТекст.SelectionColor = Color.Black;
            panelЦвета.Visible = false;
        }

        private void label77_Click(object sender, EventArgs e)
        {
            richТекст.SelectionColor = Color.DimGray;
            panelЦвета.Visible = false;
        }

        private void label76_Click(object sender, EventArgs e)
        {
            richТекст.SelectionColor = Color.DarkGray;
            panelЦвета.Visible = false;
        }

        private void label75_Click(object sender, EventArgs e)
        {
            richТекст.SelectionColor = Color.Brown;
            panelЦвета.Visible = false;
        }

        private void label74_Click(object sender, EventArgs e)
        {
            richТекст.SelectionColor = Color.Red;
            panelЦвета.Visible = false;
        }

        private void label51_Click(object sender, EventArgs e)
        {
            richТекст.SelectionColor = Color.LightPink;
            panelЦвета.Visible = false;
        }

        private void label73_Click(object sender, EventArgs e)
        {
            richТекст.SelectionColor = Color.SaddleBrown;
            panelЦвета.Visible = false;
        }

        private void label72_Click(object sender, EventArgs e)
        {
            richТекст.SelectionColor = Color.Chocolate;
            panelЦвета.Visible = false;
        }

        private void label71_Click(object sender, EventArgs e)
        {
            richТекст.SelectionColor = Color.SandyBrown;
            panelЦвета.Visible = false;
        }

        private void label70_Click(object sender, EventArgs e)
        {
            richТекст.SelectionColor = Color.DarkOrange;
            panelЦвета.Visible = false;
        }

        private void label69_Click(object sender, EventArgs e)
        {
            richТекст.SelectionColor = Color.Gold;
            panelЦвета.Visible = false;
        }

        private void label68_Click(object sender, EventArgs e)
        {
            richТекст.SelectionColor = Color.Yellow;
            panelЦвета.Visible = false;
        }

        private void label50_Click(object sender, EventArgs e)
        {
            richТекст.SelectionColor = Color.Olive;
            panelЦвета.Visible = false;
        }

        private void label67_Click(object sender, EventArgs e)
        {
            richТекст.SelectionColor = Color.DarkKhaki;
            panelЦвета.Visible = false;
        }

        private void label66_Click(object sender, EventArgs e)
        {
            richТекст.SelectionColor = Color.Khaki;
            panelЦвета.Visible = false;
        }

        private void label65_Click(object sender, EventArgs e)
        {
            richТекст.SelectionColor = Color.DarkOliveGreen;
            panelЦвета.Visible = false;
        }

        private void label64_Click(object sender, EventArgs e)
        {
            richТекст.SelectionColor = Color.Green;
            panelЦвета.Visible = false;
        }

        private void label63_Click(object sender, EventArgs e)
        {
            richТекст.SelectionColor = Color.LimeGreen;
            panelЦвета.Visible = false;
        }

        private void label62_Click(object sender, EventArgs e)
        {
            richТекст.SelectionColor = Color.SeaGreen;
            panelЦвета.Visible = false;
        }

        private void label49_Click(object sender, EventArgs e)
        {
            richТекст.SelectionColor = Color.MediumSeaGreen;
            panelЦвета.Visible = false;
        }

        private void label61_Click(object sender, EventArgs e)
        {
            richТекст.SelectionColor = Color.SpringGreen;
            panelЦвета.Visible = false;
        }

        private void label60_Click(object sender, EventArgs e)
        {
            richТекст.SelectionColor = Color.DarkSlateGray;
            panelЦвета.Visible = false;
        }

        private void label59_Click(object sender, EventArgs e)
        {
            richТекст.SelectionColor = Color.LightSeaGreen;
            panelЦвета.Visible = false;
        }

        private void label58_Click(object sender, EventArgs e)
        {
            richТекст.SelectionColor = Color.Aquamarine;
            panelЦвета.Visible = false;
        }

        private void label57_Click(object sender, EventArgs e)
        {
            richТекст.SelectionColor = Color.MidnightBlue;
            panelЦвета.Visible = false;
        }

        private void label56_Click(object sender, EventArgs e)
        {
            richТекст.SelectionColor = Color.Blue;
            panelЦвета.Visible = false;
        }

        private void label48_Click(object sender, EventArgs e)
        {
            richТекст.SelectionColor = Color.CornflowerBlue;
            panelЦвета.Visible = false;
        }

        private void label55_Click(object sender, EventArgs e)
        {
            richТекст.SelectionColor = Color.LightSteelBlue;
            panelЦвета.Visible = false;
        }

        private void label54_Click(object sender, EventArgs e)
        {
            richТекст.SelectionColor = Color.DarkMagenta;
            panelЦвета.Visible = false;
        }

        private void label53_Click(object sender, EventArgs e)
        {
            richТекст.SelectionColor = Color.DarkViolet;
            panelЦвета.Visible = false;
        }

        private void label52_Click(object sender, EventArgs e)
        {
            richТекст.SelectionColor = Color.Thistle;
            panelЦвета.Visible = false;
        }

        private void label46_Click(object sender, EventArgs e)
        {
            richТекст.SelectionColor = Color.MediumVioletRed;
            panelЦвета.Visible = false;
        }

        private void label45_Click(object sender, EventArgs e)
        {
            richТекст.SelectionColor = Color.DeepPink;
            panelЦвета.Visible = false;
        }

        private void label47_Click(object sender, EventArgs e)
        {
            richТекст.SelectionColor = Color.HotPink;
            panelЦвета.Visible = false;
        }

        private void btnНайтиПисьмо_Click(object sender, EventArgs e)
        {
            try
            {
                for (int i = 0; i < dataGridView.RowCount; i++)
                {
                    dataGridView.Rows[i].Selected = false;
                    for (int j = 0; j < dataGridView.ColumnCount; j++)
                        if (dataGridView.Rows[i].Cells[j].Value != null)
                            if (dataGridView.Rows[i].Cells[j].Value.ToString().Contains(cueПоискПисем.Text))
                            {
                                dataGridView.Rows[i].Selected = true;
                                break;
                            }
                }
            }
            catch (Exception)
            {
                MessageBox.Show("Письма с такими данными не найдены!", "Сообщение", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            cueПоискПисем.Text = "";
        }
        
        private void cueСсылка_Enter(object sender, EventArgs e)
        {
            cueСсылка.BackColor = Color.Honeydew;
        }

        private void cueСсылка_Leave(object sender, EventArgs e)
        {
            cueСсылка.BackColor = Color.PaleGreen;
        }

        private void labelВложение_Click(object sender, EventArgs e)
        {
            try
            {
                string path = txtПрикрепитьФайл.Text;
                Process.Start(path);
            }
            catch (Exception)
            {
                MessageBox.Show("Данное письмо не имеет вложений!", "Сообщение", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        private void выделитьВсеToolStripMenuItem_Click(object sender, EventArgs e)
        {
            richТекст.SelectAll();
        }

        private void копироватьToolStripMenuItem_Click(object sender, EventArgs e)
        {
            richТекст.Copy();
        }

        private void вставитьToolStripMenuItem_Click(object sender, EventArgs e)
        {
            richТекст.Paste();
        }

        private void удалитьToolStripMenuItem_Click(object sender, EventArgs e)
        {
            richТекст.SelectedText = "";
        }

        private void заполнитьПисьмоДаннымиИзТаблицыToolStripMenuItem_Click(object sender, EventArgs e)
        {
            comboЗаполнить.Visible = true;
            comboЗаполнить2.Visible = true;
        }

        private void comboЗаполнить_SelectedValueChanged(object sender, EventArgs e)
        {
            richТекст.SelectedText = comboЗаполнить.Text;
        }

        private void comboЗаполнить2_SelectedValueChanged(object sender, EventArgs e)
        {
            richТекст.SelectedText = comboЗаполнить2.Text;
        }

        private void linkАвторизация2_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            panelОтправкаЕмайл.Visible = false;
            panelИнструменты.Visible = false;
            panelСохранить.Visible = false;
            dataGridView.Visible = false;
            treeEmail.Visible = false;
            panelОтправка.Visible = false;
            btnGmail.Visible = false;
            panelАвторизация.Visible = true;
        }

        private void comboЗаполнить_Enter(object sender, EventArgs e)
        {
            comboЗаполнить.BackColor = Color.Honeydew;
        }

        private void comboЗаполнить_Leave(object sender, EventArgs e)
        {
            comboЗаполнить.BackColor = Color.PaleGreen;
        }

        private void comboЗаполнить2_Enter(object sender, EventArgs e)
        {
            comboЗаполнить2.BackColor = Color.Honeydew;
        }

        private void comboЗаполнить2_Leave(object sender, EventArgs e)
        {
            comboЗаполнить2.BackColor = Color.PaleGreen;
        }
    }
}
