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
                Text = carpetaSeleccionada
            };

            // TextBox para búsqueda
            var txtBuscar = new TextBox
            {
                Location = new Point(10, 50),
                Size = new Size(200, 30),
                PlaceholderText = "Buscar...",
                Name = "txtBuscar"
            };

            // Añadir checkboxes para filtrar por estado de pago
            var chkPagado = new CheckBox
            {
                Text = "Pagados",
                Location = new Point(330, 50),
                AutoSize = true,
                Name = "chkPagado"
            };

            var chkPendiente = new CheckBox
            {
                Text = "Pendientes",
                Location = new Point(420, 50),
                AutoSize = true,
                Name = "chkPendiente",
                Checked = false // Por defecto, mostrar pendientes
            };

            var chkOcultarFechas = new CheckBox
            {
                Text = "Ocultar fechas como cliente",
                Location = new Point(530, 50),
                AutoSize = true,
                Name = "chkOcultarFechas",
                Checked = true // Por defecto, ocultar fechas
            };

            var chkSoloFechas = new CheckBox
            {
                Text = "Solo mostrar fechas como cliente",
                Location = new Point(530, 75), // Posición debajo del otro checkbox
                AutoSize = true,
                Name = "chkSoloFechas",
                Checked = false // Por defecto, no activado
            };

            chkOcultarFechas.CheckedChanged += (s, e) =>
            {
                if (chkOcultarFechas.Checked && chkSoloFechas.Checked)
                    chkSoloFechas.Checked = false;
            };

            chkSoloFechas.CheckedChanged += (s, e) =>
            {
                if (chkSoloFechas.Checked && chkOcultarFechas.Checked)
                    chkOcultarFechas.Checked = false;
            };


            chkPagado.CheckedChanged += (s, e) => RealizarBusqueda();
            chkPendiente.CheckedChanged += (s, e) => RealizarBusqueda();
            chkOcultarFechas.CheckedChanged += (s, e) => RealizarBusqueda();
            chkSoloFechas.CheckedChanged += (s, e) => RealizarBusqueda();

            chkOcultarFechas.CheckedChanged += (s, e) =>
            {
                if (chkOcultarFechas.Checked && chkSoloFechas.Checked)
                {
                    chkSoloFechas.Checked = false;
                    // No necesitamos llamar a RealizarBusqueda() aquí porque el evento
                    // CheckedChanged de chkSoloFechas ya lo hará
                }
            };

            chkSoloFechas.CheckedChanged += (s, e) =>
            {
                if (chkSoloFechas.Checked && chkOcultarFechas.Checked)
                {
                    chkOcultarFechas.Checked = false;
                    // No necesitamos llamar a RealizarBusqueda() aquí porque el evento
                    // CheckedChanged de chkOcultarFechas ya lo hará
                }
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
            panelBusqueda.Controls.AddRange(new Control[] {
    btnCarpeta, txtRuta, txtBuscar, btnBuscar,
    chkPagado, chkPendiente, chkOcultarFechas, chkSoloFechas
});
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


        private void RealizarBusqueda()
        {
            if (string.IsNullOrEmpty(carpetaSeleccionada)) return;

            var txtBuscar = Controls.Find("txtBuscar", true).FirstOrDefault() as TextBox;
            var dgvResultados = Controls.Find("dgvResultados", true).FirstOrDefault() as DataGridView;
            var chkPagado = Controls.Find("chkPagado", true).FirstOrDefault() as CheckBox;
            var chkPendiente = Controls.Find("chkPendiente", true).FirstOrDefault() as CheckBox;
            var chkOcultarFechas = Controls.Find("chkOcultarFechas", true).FirstOrDefault() as CheckBox;
            var chkSoloFechas = Controls.Find("chkSoloFechas", true).FirstOrDefault() as CheckBox;

            if (txtBuscar == null || dgvResultados == null ||
                chkPagado == null || chkPendiente == null ||
                chkOcultarFechas == null || chkSoloFechas == null)
                return;

            // Verificar que al menos un filtro de estado esté seleccionado
            if (!chkPagado.Checked && !chkPendiente.Checked)
            {
                // No mostrar mensajes aquí para evitar spam
                // Solo establecer un estado predeterminado
                chkPendiente.Checked = true;
            }

            BuscarEnArchivos(
                txtBuscar.Text.ToLower(),
                dgvResultados,
                chkPagado.Checked,
                chkPendiente.Checked,
                chkOcultarFechas.Checked,
                chkSoloFechas.Checked);
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
            var chkPagado = Controls.Find("chkPagado", true).FirstOrDefault() as CheckBox;
            var chkPendiente = Controls.Find("chkPendiente", true).FirstOrDefault() as CheckBox;
            var chkOcultarFechas = Controls.Find("chkOcultarFechas", true).FirstOrDefault() as CheckBox;
            var chkSoloFechas = Controls.Find("chkSoloFechas", true).FirstOrDefault() as CheckBox;

            if (txtBuscar != null && dgvResultados != null &&
                chkPagado != null && chkPendiente != null &&
                chkOcultarFechas != null && chkSoloFechas != null)
            {
                // Eliminar la validación que requería al menos un checkbox seleccionado

                BuscarEnArchivos(
                    txtBuscar.Text.ToLower(),
                    dgvResultados,
                    chkPagado.Checked,
                    chkPendiente.Checked,
                    chkOcultarFechas.Checked,
                    chkSoloFechas.Checked);
            }
        }
        private void BuscarEnArchivos(string textoBusqueda, DataGridView dgv, bool mostrarPagados,
                     bool mostrarPendientes, bool ocultarFechas, bool soloFechas)
        {
            dgv.Rows.Clear();
            if (!Directory.Exists(carpetaSeleccionada)) return;

            // Determinar si se deben mostrar todos (cuando ninguno está seleccionado)
            bool mostrarTodos = !mostrarPagados && !mostrarPendientes;

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
                // Aplicar filtro por estado de pago - nueva lógica
                if (!mostrarTodos && ((resultado.estaPagado && !mostrarPagados) ||
                    (!resultado.estaPagado && !mostrarPendientes)))
                {
                    continue; // Saltar este resultado si no cumple con el filtro
                }

                bool esFechaCliente = EsFormatoFecha(resultado.cliente);

                // Aplicar filtros de fecha
                if ((ocultarFechas && esFechaCliente) || (soloFechas && !esFechaCliente))
                {
                    continue; // Saltar según los filtros de fecha
                }

                int rowIndex = dgv.Rows.Add(
                    resultado.nombreCompleto,
                    resultado.cliente,
                    $"{resultado.articulosActivos} artículos: {resultado.detallesArticulos}",
                    $"${resultado.total:N2} (Com: ${resultado.comision:N2})",
                    resultado.estaPagado ? "PAGADO" : "PENDIENTE",
                    "Abrir"
                );

                // Colorear la celda de estado
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

        // Función para determinar si un string tiene formato de fecha
        private bool EsFormatoFecha(string texto)
        {
            // Verificar si contiene formatos de fecha comunes
            return texto.Contains("/20") || // Busca patrones como "10/11/2024"
                   texto.Contains("0:00:00") || // Busca horas como "0:00:00"
                   DateTime.TryParse(texto, out _); // Intenta parsear como fecha
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