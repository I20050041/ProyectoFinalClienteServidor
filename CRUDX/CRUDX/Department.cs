using Microsoft.VisualBasic;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Net;
using System.Security.Cryptography.Pkcs;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Xml.Linq;
using static System.Net.Mime.MediaTypeNames;
using static System.Windows.Forms.VisualStyles.VisualStyleElement;

namespace CRUD
{
    public partial class Department : Form
    {
        public Department()
        {
            InitializeComponent();
        }
        private const string connectionString = "Data Source=DESKTOP-KB3P8F9\\SQLEXPRESS;Initial Catalog=AreaAcademicaBn;Integrated Security=True;"; // Reemplaza con tu cadena de conexión

        private void AcademicAward_Load(object sender, EventArgs e)
        {
            dateTimePicker1.Format = DateTimePickerFormat.Custom;

            dateTimePicker1.CustomFormat = "MM/dd/yyyy";
            LoadAcademicAward();
        }
        private void dataGridView1_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            LoadDataFromGrid();
        }

        private void LoadAcademicAward()
        {
            using (SqlConnection connection = new SqlConnection(connectionString))
            {
                connection.Open();
                string query = "SELECT \r\n    Department.idDepartment,\r\n    Department.name AS DepartmentName,\r\n    Department.description,\r\n    Department.director,\r\n    Department.foundingDate,\r\n    Department.phone,\r\n    Department.email,\r\n    Department.location,\r\n    Department.openingHour,\r\n    Department.closingHour,\r\n    Department.status AS DepartmentStatus,\r\n    Employee.idEmployee,\r\n    Employee.name AS EmployeeName,\r\n    Employee.lastName AS EmployeeLastName\r\nFROM \r\n    Department\r\nLEFT JOIN \r\n    Employee ON Department.idEmployee = Employee.idEmployee\r\nWHERE \r\n    Department.status = 1;";
                SqlDataAdapter adapter = new SqlDataAdapter(query, connection);
                DataTable dataTable = new DataTable();
                adapter.Fill(dataTable);
                dataGridView1.DataSource = dataTable;
            }
            using (SqlConnection connection = new SqlConnection(connectionString))
            {
                connection.Open();
                string query = "SELECT name, idEmployee FROM Employee WHERE status =1";
                SqlDataAdapter adapter = new SqlDataAdapter(query, connection);
                DataTable combotable = new DataTable();
                comboBox1.Items.Clear();
                adapter.Fill(combotable);
                foreach (DataRow row in combotable.Rows)
                {
                    string supplierInfo = $"{row["idEmployee"]} - {row["name"]}";
                    comboBox1.Items.Add(supplierInfo);
                }
            }
        }
        private void button1_Click(object sender, EventArgs e)
        {
            string ReportInfo = comboBox1.SelectedItem.ToString();
            int idStudent = 0;

            // Obtener el EmployeeID de ReportInfo
            string[] reportInfoParts = ReportInfo.Split('-');
            if (reportInfoParts.Length >= 2)
            {
                int.TryParse(reportInfoParts[0].Trim(), out idStudent);
            }
            using (SqlConnection connection = new SqlConnection(connectionString))
            {
                connection.Open();
                string query = "INSERT INTO Department (name, description, director, foundingDate, phone, email, location, openingHour, closingHour, status, idEmployee) " +
               "VALUES (@name, @description, @director, @foundingDate, @phone, @email, @location, @openingHour, @closingHour, @status, @idEmployee)";
                SqlCommand cmd = new SqlCommand(query, connection);
                cmd.Parameters.AddWithValue("@name", txtName.Text);
                cmd.Parameters.AddWithValue("@description", txtDescription.Text);
                cmd.Parameters.AddWithValue("@director", txtDirector.Text);
                cmd.Parameters.AddWithValue("@foundingDate", dateTimePicker1.Text);
                cmd.Parameters.AddWithValue("@phone", txtPhone.Text);
                cmd.Parameters.AddWithValue("@email", txtEmail.Text);
                cmd.Parameters.AddWithValue("@location", txtLocation.Text);
                cmd.Parameters.AddWithValue("@openingHour", txtOpeningHour.Text);
                cmd.Parameters.AddWithValue("@closingHour", txtClosingHour.Text);
                cmd.Parameters.AddWithValue("@status", chkStatus.Checked);
                cmd.Parameters.AddWithValue("@idEmployee", idStudent);
                cmd.ExecuteNonQuery();
                LoadAcademicAward();
            }
        }
        private void button2_Click(object sender, EventArgs e)
        {
            if (dataGridView1.SelectedRows.Count > 0)
            {
                int departmentId = Convert.ToInt32(dataGridView1.SelectedRows[0].Cells["idDepartment"].Value);

                string ReportInfo = comboBox1.SelectedItem.ToString();
                int idStudent = 0;

                // Obtener el EmployeeID de ReportInfo
                string[] reportInfoParts = ReportInfo.Split('-');
                if (reportInfoParts.Length >= 2)
                {
                    int.TryParse(reportInfoParts[0].Trim(), out idStudent);
                }
                using (SqlConnection connection = new SqlConnection(connectionString))
                {
                    connection.Open();
                    string query = "UPDATE Department " +
                       "SET name = @name, description = @description, director = @director, " +
                       "foundingDate = @foundingDate, phone = @phone, email = @email, " +
                       "location = @location, openingHour = @openingHour, closingHour = @closingHour, " +
                       "status = @status, idEmployee = @idEmployee " +
                       "WHERE idDepartment = @idDepartment";
                    SqlCommand cmd = new SqlCommand(query, connection);
                    cmd.Parameters.AddWithValue("@idDepartment", departmentId);
                    cmd.Parameters.AddWithValue("@name", txtName.Text);
                    cmd.Parameters.AddWithValue("@description", txtDescription.Text);
                    cmd.Parameters.AddWithValue("@director", txtDirector.Text);
                    cmd.Parameters.AddWithValue("@foundingDate", dateTimePicker1.Text);
                    cmd.Parameters.AddWithValue("@phone", txtPhone.Text);
                    cmd.Parameters.AddWithValue("@email", txtEmail.Text);
                    cmd.Parameters.AddWithValue("@location", txtLocation.Text);
                    cmd.Parameters.AddWithValue("@openingHour", txtOpeningHour.Text);
                    cmd.Parameters.AddWithValue("@closingHour", txtClosingHour.Text);
                    cmd.Parameters.AddWithValue("@status", chkStatus.Checked);
                    cmd.Parameters.AddWithValue("@idEmployee", idStudent);
                    cmd.ExecuteNonQuery();
                    LoadAcademicAward();
                }
            }
            else
            {
                MessageBox.Show("Seleccione un departamento para actualizar.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        private void LoadDataFromGrid()
        {
            if (dataGridView1.SelectedRows.Count > 0)
            {
                DataGridViewRow selectedRow = dataGridView1.SelectedRows[0];
                txtName.Text = selectedRow.Cells["DepartmentName"].Value.ToString();
                txtDescription.Text = selectedRow.Cells["description"].Value.ToString();
                txtDirector.Text = selectedRow.Cells["director"].Value.ToString();
                dateTimePicker1.Text = selectedRow.Cells["foundingDate"].Value.ToString();
                txtPhone.Text = selectedRow.Cells["phone"].Value.ToString();
                txtEmail.Text = selectedRow.Cells["email"].Value.ToString();
                txtLocation.Text = selectedRow.Cells["location"].Value.ToString();
                txtOpeningHour.Text = selectedRow.Cells["openingHour"].Value.ToString();
                txtClosingHour.Text = selectedRow.Cells["closingHour"].Value.ToString();
                comboBox1.SelectedValue = selectedRow.Cells["idEmployee"].Value;
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            if (dataGridView1.SelectedRows.Count > 0)
            {
                int departmentId = Convert.ToInt32(dataGridView1.SelectedRows[0].Cells["idDepartment"].Value);

                using (SqlConnection connection = new SqlConnection(connectionString))
                {
                    connection.Open();
                    string query = "Update Department set status = 0 WHERE idDepartment = @idDepartment";

                    SqlCommand cmd = new SqlCommand(query, connection);
                    cmd.Parameters.AddWithValue("@idDepartment", departmentId);

                    cmd.ExecuteNonQuery();
                    LoadAcademicAward();
                }
            }
            else
            {
                MessageBox.Show("Seleccione un departamento para eliminar.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
    }
}
