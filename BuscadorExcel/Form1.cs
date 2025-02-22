using System;
using System.Windows.Forms;
using System.IO;
using ClosedXML.Excel;
using System.Collections.Generic;
using System.Linq;

namespace BuscadorExcel
{
    public partial class Form1 : Form
    {
        private string carpetaSeleccionada = string.Empty;

        public Form1()
        {
            InitializeComponent();
            ConfigurarFormulario();
        }

        private void ConfigurarFormulario()
        {
            this.Text = "Buscador de Pedidos Excel";
            this.WindowState = FormWindowState.Maximized;

            // Panel superior para controles de búsqueda
            var panelBusqueda = new Panel
            {
                Dock = DockStyle.Top,
                Height = 100
            };

            // Botón para seleccionar carpeta
            var btnCarpeta = new Button
            {
                Text = "Seleccionar Carpeta",
                Location = new Point(10, 10),
                Size = new Size(150, 30),
                Name = "btnCarpeta"
            };
            btnCarpeta.Click += BtnCarpeta_Click!;

            // TextBox para mostrar la ruta
            var txtRuta = new TextBox
            {
                Location = new Point(170, 15),
                Size = new Size(400, 30),
                ReadOnly = true,
                Name = "txtRuta"
            };

            // TextBox para búsqueda
            var txtBuscar = new TextBox
            {
                Location = new Point(10, 50),
                Size = new Size(200, 30),
                PlaceholderText = "Buscar...",
                Name = "txtBuscar"
            };

            // Botón buscar
            var btnBuscar = new Button
            {
                Text = "Buscar",
                Location = new Point(220, 50),
                Size = new Size(100, 30),
                Name = "btnBuscar"
            };
            btnBuscar.Click += BtnBuscar_Click!;

            // DataGridView para resultados
            var dgvResultados = new DataGridView
            {
                Dock = DockStyle.Fill,
                AllowUserToAddRows = false,
                ReadOnly = true,
                AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill,
                Name = "dgvResultados"
            };

            // Configurar columnas del DataGridView
            dgvResultados.Columns.AddRange(new DataGridViewColumn[]
            {
                new DataGridViewTextBoxColumn { Name = "Archivo", HeaderText = "Fecha (Archivo)", Width = 100 },
                new DataGridViewTextBoxColumn { Name = "Cliente", HeaderText = "Cliente", Width = 150 },
                new DataGridViewTextBoxColumn { Name = "Articulos", HeaderText = "Artículos", Width = 200 },
                new DataGridViewTextBoxColumn { Name = "Total", HeaderText = "Total", Width = 100 },
                new DataGridViewLinkColumn { Name = "Abrir", HeaderText = "Abrir Excel", Width = 80 }
            });

            dgvResultados.CellClick += DgvResultados_CellClick!;

            // Agregar controles
            panelBusqueda.Controls.AddRange(new Control[] { btnCarpeta, txtRuta, txtBuscar, btnBuscar });
            this.Controls.Add(dgvResultados);
            this.Controls.Add(panelBusqueda);
        }

        private void BtnCarpeta_Click(object sender, EventArgs e)
        {
            using var fbd = new FolderBrowserDialog();
            if (fbd.ShowDialog() == DialogResult.OK)
            {
                carpetaSeleccionada = fbd.SelectedPath;
                var txtRuta = Controls.Find("txtRuta", true).FirstOrDefault() as TextBox;
                if (txtRuta != null)
                {
                    txtRuta.Text = carpetaSeleccionada;
                }
            }
        }

        private void BtnBuscar_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrEmpty(carpetaSeleccionada))
            {
                MessageBox.Show("Por favor, seleccione una carpeta primero", "Error", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            var txtBuscar = Controls.Find("txtBuscar", true).FirstOrDefault() as TextBox;
            var dgvResultados = Controls.Find("dgvResultados", true).FirstOrDefault() as DataGridView;
            
            if (txtBuscar != null && dgvResultados != null)
            {
                BuscarEnArchivos(txtBuscar.Text.ToLower(), dgvResultados);
            }
        }

        private void BuscarEnArchivos(string textoBusqueda, DataGridView dgv)
{
    dgv.Rows.Clear();
    if (!Directory.Exists(carpetaSeleccionada)) return;

    foreach (var archivo in Directory.GetFiles(carpetaSeleccionada, "*.xlsx"))
    {
        try
        {
            using (var workbook = new XLWorkbook(new FileStream(archivo, FileMode.Open, FileAccess.Read, FileShare.ReadWrite)))
            {
                foreach (var worksheet in workbook.Worksheets)
                {
                    string nombreHoja = worksheet.Name.ToLower();
                    var clienteCell = worksheet.Cell("B1");
                    
                    string cliente = !clienteCell.IsEmpty() 
                        ? clienteCell.GetString().ToLower() 
                        : nombreHoja;

                    if (cliente.Contains(textoBusqueda))
                    {
                        var articulosActivos = 0;
                        var total = 0.0M;
                        var detallesArticulos = new List<string>();

                        var row = 3;  // Empezar después de los encabezados
                        while (!worksheet.Cell($"B{row}").IsEmpty())
                        {
                            try
                            {
                                var estadoCell = worksheet.Cell($"E{row}");
                                if (!estadoCell.IsEmpty())
                                {
                                    string estadoStr = estadoCell.GetString().ToLower();
                                    bool estaSeleccionado = !estadoStr.Contains("falso") && 
                                                          !estadoStr.Contains("false") &&
                                                          estadoStr != "0" &&
                                                          !string.IsNullOrEmpty(estadoStr);

                                    if (estaSeleccionado)
                                    {
                                        articulosActivos++;
                                        
                                        var precioCell = worksheet.Cell($"F{row}");
                                        if (!precioCell.IsEmpty())
                                        {
                                            try
{
    // Intenta obtener el valor como número primero
    if (precioCell.Value.IsNumber)
    {
        total += (decimal)precioCell.Value.GetNumber();
    }
    else
    {
        // Si no es número, intenta parsear el string
        string precioStr = precioCell.GetString()
            .Replace("$", "")
            .Replace(",", ".")
            .Trim();
                                                    
        if (decimal.TryParse(precioStr, out decimal precio))
        {
            total += precio;
        }
    }
}
catch
{
    // Si falla la conversión, intenta un último método
    try
    {
        total += (decimal)precioCell.GetDouble();
    }
    catch
    {
        // Si todo falla, ignora este precio
        Console.WriteLine($"No se pudo leer el precio en la fila {row}");
    }
}
                                        }

                                        var articuloCell = worksheet.Cell($"D{row}");
                                        if (!articuloCell.IsEmpty())
                                        {
                                            detallesArticulos.Add(articuloCell.GetString());
                                        }
                                    }
                                }
                            }
                            catch (Exception ex)
                            {
                                Console.WriteLine($"Error en fila {row}: {ex.Message}");
                            }
                            row++;
                        }

                        decimal comision = articulosActivos * 0.50M;
                        decimal totalConComision = total + comision;

                        string nombreArchivo = Path.GetFileNameWithoutExtension(archivo);
                        string articulosStr = detallesArticulos.Count > 0 ? 
                            string.Join(", ", detallesArticulos) : "";

                        dgv.Rows.Add(
                            nombreArchivo,
                            cliente,
                            $"{articulosActivos} artículos: {articulosStr}",
                            $"${total:N2} (Com: ${comision:N2})",
                            "Abrir"
                        );
                    }
                }
            }
        }
        catch (Exception ex)
        {
            // Solo mostramos el error si no está relacionado con imágenes
            if (!ex.Message.Contains("Picture names"))
            {
                MessageBox.Show(
                    $"Error al leer el archivo {Path.GetFileName(archivo)}: {ex.Message}",
                    "Error",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error
                );
            }
            // Si es un error de imágenes, lo ignoramos y continuamos
            continue;
        }
    }

    if (dgv.Rows.Count == 0)
    {
        MessageBox.Show(
            $"No se encontraron resultados para '{textoBusqueda}'", 
            "Sin resultados", 
            MessageBoxButtons.OK, 
            MessageBoxIcon.Information
        );
    }
}

        private void DgvResultados_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            var dgv = sender as DataGridView;
            if (dgv == null) return;

            if (e.ColumnIndex == dgv.Columns["Abrir"].Index && e.RowIndex >= 0)
            {
                string nombreArchivo = dgv.Rows[e.RowIndex].Cells["Archivo"].Value?.ToString() ?? "";
                if (string.IsNullOrEmpty(nombreArchivo)) return;

                string rutaCompleta = Path.Combine(carpetaSeleccionada, nombreArchivo + ".xlsx");
                
                try
                {
                    System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo
                    {
                        FileName = rutaCompleta,
                        UseShellExecute = true
                    });
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"Error al abrir el archivo: {ex.Message}");
                }
            }
        }
    }
}