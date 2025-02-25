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
        private string carpetaSeleccionada = @"C:\Users\ASUS\Desktop\UniversidadZzzz\Facturas";

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
                Name = "txtRuta",
                Text = carpetaSeleccionada  // Añade esta línea
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
    new DataGridViewTextBoxColumn { Name = "EstadoPago", HeaderText = "Estado de Pago", Width = 100 },
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

            // Lista para almacenar todos los resultados antes de ordenarlos
            var resultados = new List<(string nombreCompleto, string cliente, int articulosActivos,
                                       string detallesArticulos, decimal total, decimal comision, bool estaPagado)>();

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
                                bool estaPagado = false;

                                // Verificar si está pagado - primero buscar en la celda específica F3
                                try
                                {
                                    var pagadoCell = worksheet.Cell("F3");
                                    if (!pagadoCell.IsEmpty())
                                    {
                                        string estadoPago = pagadoCell.GetString().ToUpper();
                                        if (estadoPago.Contains("PAGADO") && !estadoPago.Contains("NO PAGADO"))
                                        {
                                            estaPagado = true;
                                        }
                                    }
                                }
                                catch { }

                                // Si no encontramos en F3, buscar en todas las celdas F
                                if (!estaPagado)
                                {
                                    try
                                    {
                                        for (int r = 2; r <= 20; r++)
                                        {
                                            var cell = worksheet.Cell(r, 6); // Columna F (índice 6)
                                            if (!cell.IsEmpty())
                                            {
                                                string valor = cell.GetString().ToUpper();
                                                if (valor.Contains("PAGADO") && !valor.Contains("NO PAGADO"))
                                                {
                                                    estaPagado = true;
                                                    break;
                                                }
                                            }
                                        }
                                    }
                                    catch { }
                                }

                                // Si aún no encontramos, buscar en cualquier parte de la hoja
                                if (!estaPagado)
                                {
                                    try
                                    {
                                        for (int r = 1; r <= 20; r++)
                                        {
                                            for (int c = 1; c <= 10; c++)
                                            {
                                                var cell = worksheet.Cell(r, c);
                                                if (!cell.IsEmpty())
                                                {
                                                    string valor = cell.GetString().ToUpper();
                                                    if (valor.Contains("PAGADO") && !valor.Contains("NO PAGADO"))
                                                    {
                                                        estaPagado = true;
                                                        break;
                                                    }
                                                }
                                            }
                                            if (estaPagado) break;
                                        }
                                    }
                                    catch { }
                                }

                                var fila = 3;  // Empezar después de los encabezados
                                while (!worksheet.Cell($"B{fila}").IsEmpty())
                                {
                                    try
                                    {
                                        var estadoCell = worksheet.Cell($"E{fila}");
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

                                                var precioCell = worksheet.Cell($"F{fila}");
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
                                                        }
                                                    }
                                                }

                                                var articuloCell = worksheet.Cell($"D{fila}");
                                                if (!articuloCell.IsEmpty())
                                                {
                                                    detallesArticulos.Add(articuloCell.GetString());
                                                }
                                            }
                                        }
                                    }
                                    catch { }
                                    fila++;
                                }

                                decimal comision = articulosActivos * 0.50M;
                                string nombreArchivo = Path.GetFileNameWithoutExtension(archivo);
                                string detallesStr = string.Join(", ", detallesArticulos);

                                // Guardar resultados para ordenar después
                                resultados.Add((nombreArchivo, cliente, articulosActivos, detallesStr, total, comision, estaPagado));
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
                }
            }

            // Ordenar resultados por número de archivo
            var resultadosOrdenados = resultados
                .OrderBy(r => ObtenerNumeroArchivo(r.nombreCompleto))
                .ToList();

            // Añadir resultados ordenados al DataGridView
            foreach (var resultado in resultadosOrdenados)
            {
                int rowIndex = dgv.Rows.Add(
                    resultado.nombreCompleto,
                    resultado.cliente,
                    $"{resultado.articulosActivos} artículos: {resultado.detallesArticulos}",
                    $"${resultado.total:N2} (Com: ${resultado.comision:N2})",
                    resultado.estaPagado ? "PAGADO" : "PENDIENTE",
                    "Abrir"
                );

                // Colorear la celda de estado usando el índice en lugar del nombre
                dgv.Rows[rowIndex].Cells[4].Style.ForeColor = resultado.estaPagado ?
                    Color.Green : Color.Red;
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

        // Función para extraer el número de archivo del formato "(N) fecha"
        private int ObtenerNumeroArchivo(string nombreArchivo)
        {
            try
            {
                // Buscar un patrón como "(1)" o "(10)" al principio del nombre
                var match = System.Text.RegularExpressions.Regex.Match(nombreArchivo, @"^\((\d+)\)");
                if (match.Success && match.Groups.Count > 1)
                {
                    return int.Parse(match.Groups[1].Value);
                }
            }
            catch { }

            // Si no se puede extraer, devolver un valor alto para que aparezca al final
            return 9999;
        }

        // Método para configurar las columnas del DataGridView
        private void ConfigurarDataGridView()
        {
            var dgvResultados = Controls.Find("dgvResultados", true).FirstOrDefault() as DataGridView;
            if (dgvResultados == null) return;

            dgvResultados.Columns.Clear();
            dgvResultados.Columns.AddRange(new DataGridViewColumn[]
            {
        new DataGridViewTextBoxColumn { Name = "Archivo", HeaderText = "Fecha (Archivo)", Width = 100 },
        new DataGridViewTextBoxColumn { Name = "Cliente", HeaderText = "Cliente", Width = 150 },
        new DataGridViewTextBoxColumn { Name = "Articulos", HeaderText = "Artículos", Width = 200 },
        new DataGridViewTextBoxColumn { Name = "Total", HeaderText = "Total", Width = 100 },
        new DataGridViewTextBoxColumn { Name = "Estado", HeaderText = "Estado", Width = 80 },
        new DataGridViewLinkColumn { Name = "Abrir", HeaderText = "Abrir Excel", Width = 80 }
            });
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