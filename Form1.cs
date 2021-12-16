using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.Threading.Tasks;

namespace IGESFileParser
{
    public partial class Form1 : Form
    {
        public Form1( )
        {
            InitializeComponent( );
            DragDrop += new DragEventHandler( Form1_DragDrop );
            DragEnter += new DragEventHandler( Form1_DragEnter );
            ControlAction( ControlActions.ClearBowlSelection );

            var controls = GetControls( ControlSections.Robots );

            foreach( var control in controls )
            {
                if( control is TextBox )
                {
                    ( (TextBox)control ).TextChanged += new EventHandler( BowlQuantities_TextChanged );
                }
            }
        }

        private SqlConnection conn;
        private bool uploading = false;
        private bool skipFirst = false;
        private string connectionstring;
        private string m_FileLocation;
        private string m_Output;
        private bool m_Revalidate = false;
        private readonly List<PEM> m_PEMList = new List<PEM>( );
        private readonly Dictionary<string, int> m_SelectedQuantities = new Dictionary<string, int>( );
        private readonly List<PEMTotal> m_PemTotals = new List<PEMTotal>( );
        private enum ControlActions
        {
            ClearBowlSelection,
            EnableBowlSelection
        }
        private enum ControlSections
        {
            Main,
            Robots
        }

        private List<Control> GetControls( ControlSections section )
        {
            var collection = new List<Control>( );

            switch( section )
            {
                case ControlSections.Main:
                    foreach( Control control in Controls )
                    {
                        if( control is TextBox || control is ComboBox )
                        {
                            collection.Add( control );
                        }
                    }
                    break;

                case ControlSections.Robots:
                    foreach( Control control in Controls )
                    {
                        if( control is GroupBox )
                        {
                            foreach( Control groupControl in control.Controls )
                            {
                                if( groupControl is TextBox || groupControl is ComboBox )
                                {
                                    collection.Add( groupControl );
                                }
                            }
                        }
                    }
                    break;
            }

            return collection;
        }

        private void ControlAction( ControlActions action )
        {
            var controls = GetControls( ControlSections.Robots );

            switch( action )
            {
                case ControlActions.ClearBowlSelection:
                    foreach( var control in controls )
                    {
                        control.Enabled = false;

                        if( control is ComboBox )
                        {
                            ( (ComboBox)control ).SelectedIndex = -1;
                            ( (ComboBox)control ).Items.Clear( );
                        }
                        else if( control is TextBox )
                        {
                            ( (TextBox)control ).Clear( );
                        }
                    }

                    processButton.Enabled = false;
                    break;

                case ControlActions.EnableBowlSelection:
                    foreach( var control in controls )
                    {
                        control.Enabled = true;
                    }

                    processButton.Enabled = true;
                    break;
            }
        }


        private bool ValidFields( )
        {
            var valid = true;
            var clearedOutput = false;
            m_SelectedQuantities.Clear( );
            var quantityCount = new Dictionary<string, int>( );

            if( !m_Revalidate )
            {
                m_Output = outputRichTextBox.Text;
            }

            var controls = GetControls( ControlSections.Main );

            foreach( var control in controls )
            {
                if( control is TextBox )
                {
                    if( string.IsNullOrEmpty( ( (TextBox)control ).Text ) || string.IsNullOrWhiteSpace( ( (TextBox)control ).Text ) )
                    {
                        if( !clearedOutput )
                        {
                            outputRichTextBox.Clear( );
                            clearedOutput = true;
                        }
                        outputRichTextBox.AppendText( $"{ ( (TextBox)control ).Tag } is empty \r\n" );
                        valid = false;
                    }
                }
                else if( control is ComboBox )
                {
                    if( ( (ComboBox)control ).SelectedIndex == -1 )
                    {
                        if( !clearedOutput )
                        {
                            outputRichTextBox.Clear( );
                            clearedOutput = true;
                        }
                        outputRichTextBox.AppendText( $"{ ( (ComboBox)control ).Tag } is unselected \r\n" );
                        valid = false;
                    }
                }
            }

            if( valid )
            {
                foreach( Control control in Controls )
                {
                    if( control is GroupBox )
                    {
                        foreach( Control groupControl in control.Controls )
                        {
                            if( groupControl is ComboBox selectedItem )
                            {
                                if( selectedItem.SelectedIndex != -1 )
                                {
                                    foreach( Control innerGroupControl in control.Controls )
                                    {
                                        if( innerGroupControl is TextBox )
                                        {
                                            if( selectedItem.Tag == innerGroupControl.Tag )
                                            {
                                                if( string.IsNullOrEmpty( innerGroupControl.Text ) || !int.TryParse( innerGroupControl.Text, out _ ) )
                                                {
                                                    if( !clearedOutput )
                                                    {
                                                        outputRichTextBox.Clear( );
                                                        clearedOutput = true;
                                                    }

                                                    outputRichTextBox.AppendText( $"Quantity for bowl { ( (TextBox)innerGroupControl ).Tag } is not set or invalid \r\n" );
                                                    valid = false;
                                                }
                                                else
                                                {
                                                    if( !m_Revalidate )
                                                    {
                                                        m_SelectedQuantities.Add( selectedItem.Tag.ToString( ), int.Parse( innerGroupControl.Text ) );

                                                        if( !quantityCount.ContainsKey( selectedItem.Text ) )
                                                        {
                                                            quantityCount.Add( selectedItem.Text, int.Parse( innerGroupControl.Text ) );
                                                        }
                                                        else
                                                        {
                                                            quantityCount[ selectedItem.Text ] += int.Parse( innerGroupControl.Text );
                                                        }
                                                    }
                                                }
                                            }
                                        }
                                    }
                                }
                            }
                        }
                    }
                }
            }

            if( valid )
            {
                foreach( var enteredTotal in quantityCount )
                {
                    foreach( var loadedTotal in m_PemTotals )
                    {
                        if( enteredTotal.Key == loadedTotal.Name )
                        {
                            if( enteredTotal.Value > loadedTotal.Total )
                            {
                                if( !clearedOutput )
                                {
                                    outputRichTextBox.Clear( );
                                    clearedOutput = true;
                                }

                                outputRichTextBox.AppendText( $"Quantity for { loadedTotal.Name } exceeds total count of { loadedTotal.Total } \r\n" );
                                valid = false;
                            }
                        }
                    }
                }
            }

            return valid;
        }

        private void FileValidation( string file )
        {
            outputRichTextBox.Clear( );
            ControlAction( ControlActions.ClearBowlSelection );

            if( file.Contains( ".igs" ) )
            {
                m_FileLocation = file;
                selectedFileLabel.Text = Path.GetFileName( m_FileLocation );
                ControlAction( ControlActions.EnableBowlSelection );
                m_PEMList.Clear( );
                m_PemTotals.Clear( );
                m_SelectedQuantities.Clear( );
                ProcessFile( );
            }
            else
            {
                m_FileLocation = "";
                outputRichTextBox.AppendText( "Not an .igs file" );
                selectedFileLabel.Text = "Invalid file type";
            }
        }

        private void ProcessFile( )
        {
            var lineList = new List<Line>( );

            try
            {
                var reader = new StreamReader( m_FileLocation );
                string unprocessedLine;
                var firstLine = true;

                while( ( unprocessedLine = reader.ReadLine( ) ) != null )
                {
                    var line = Regex.Replace( unprocessedLine.Trim( ), @"\s+", ">" ).Split( '>' ).ToList( );
                    line.RemoveAt( line.Count - 1 );
                    var lineData = Regex.Replace( line[ 0 ], @".$", "" ).Split( ',' ).ToList( );
                    var lineGroup = line[ 1 ];

                    if( lineGroup.EndsWith( "P" ) )
                    {
                        if( firstLine )
                        {
                            lineList.Add( new Line { Data = lineData, Group = lineGroup, Type = lineData[ 0 ] } );
                            firstLine = false;
                        }
                        else
                        {
                            lineList.Last( ).Data.AddRange( lineData );
                        }

                        if( line[ 0 ].Contains( ";" ) )
                        {
                            firstLine = true;
                        }
                    }
                }

                foreach( var line in lineList )
                {
                    if( line.Type == "402" )
                    {
                        var name = lineList.Find( x => x.Type == "406" && line.Data.Contains( x.Group.Trim( 'P' ) ) );
                        var coordinates = lineList.Find( x => x.Type == "116" && line.Data.Contains( x.Group.Trim( 'P' ) ) );
                        var namePattern = new Regex( @"H(.*)" );
                        if( name != null && coordinates != null )
                        {
                            var match = Regex.Match( name.Data.Last( ), @"H(.*)_\d+$" );
                            if( !match.Success )
                            {
                                match = Regex.Match( name.Data.Last( ), @"H(.*)" );
                            }

                            m_PEMList.Add( new PEM
                            {
                                Name = match.Groups[ 1 ].Value,
                                Coordinates = coordinates.Data.Select( x => double.Parse( x.Replace( 'D', 'E' ) ).ToString( ) ).ToList( ).GetRange( 1, 3 )
                            } );
                        }
                    }
                }

                foreach( var x in m_PEMList )
                {
                    Console.WriteLine( x.Name );
                }

                m_PemTotals.AddRange( m_PEMList.GroupBy( x => x.Name )
                    .Select( y => new PEMTotal { Name = y.Key, Total = y.Count( ) } ).ToList( ) );

                foreach( var pem in m_PemTotals )
                {
                    outputRichTextBox.AppendText( $"{pem.Name}: {pem.Total} \r\n" );

                    foreach( Control control in Controls )
                    {
                        if( control is GroupBox )
                        {
                            foreach( Control groupControl in control.Controls )
                            {
                                if( groupControl is ComboBox )
                                {
                                    ( (ComboBox)groupControl ).Items.Add( pem.Name );
                                }
                            }
                        }
                    }
                }
            }
            catch( IOException exception )
            {
                outputRichTextBox.Clear( );
                MessageBox.Show( $"The file could not be read: \r\n {exception.Message}" );
            }
        }

        // Deprecated file creation. CSV export is no longer needed when going to SQL Database
        private void CreateOutputFile( )
        {
            var cartTypeRegexPattern = new Regex( @"(\w{1}\d{2}\w{1}\d{2})(\w{1})(\d{1,2})" );
            var infeedCart = cartTypeRegexPattern.Match( infeedCartComboBox.Text );
            var outfeedCart = cartTypeRegexPattern.Match( outfeedCartComboBox.Text );

            var fileOutputRows = new List<string>( ) {
                $"Barcode Number,{barcodeNumberTextBox.Text},,,,,,",
                $"Cell 1 Haeger Program,{cellHaegerProgramOneTextBox.Text},,,,,,",
                $"Cell 2 Haeger Program,{cellHaegerProgramTwoTextBox.Text},,,,,,",
                $"X-Dimension (in),{xDimensionTextBox.Text},,,,,,",
                $"Y-Dimension (in),{yDimensionTextBox.Text},,,,,,",
                $"Z-Dimension (in),{zDimensionTextBox.Text},,,,,,",
                $"Flange Direction,{flangeDirectionComboBox.Text.Last( )},,,,,,",
                $"Length Opposite Flange (mm),{pemLengthOppositeFlangeTextBox.Text},,,,,,",
                $"Infeed Cart Orientation:,{(infeedCartOrientation.Text == "P" ? "Portrait" : "Landscape")},,,,,,",
                $"Outfeed Cart Orientation:,{(outfeedCartOrientation.Text == "P" ? "Portrait" : "Landscape")},,,,,,",
                $"Infeed Cart Slots:,{infeedCartSlots.Text},,,,,,",
                $"Outfeed Cart Slots:,{outfeedCartSlots.Text},,,,,,",
                $"Infeed Cart Type:,{infeedCartComboBox.Text},,,,,,",
                $"Outfeed Cart Type:,{outfeedCartComboBox.Text},,,,,,",
                $"X-Edge Check Position 1 (in),{xEdgeCheckPosition1TextBox.Text},,,,,,",
                $"X-Edge Check Position 2 (in),{xEdgeCheckPosition2TextBox.Text},,,,,,",
                $"Y-Edge Check Position (in),{yEdgeCheckPosition.Text},,,,,,",
                $"Tool Size,{toolSizeComboBox.Text.Last( )},,,,,,",
                $"Recipe Type,{recipeTypeComboBox.Text.Last( )},,,,,,",
                $"Edge Type,{edgeTypeComboBox.Text.Last( )},,,,,,",
                $"Bowl 1,{bowlOneComboBox.Text},,{( string.IsNullOrEmpty(bowlOneQuantity.Text) ? string.Empty : bowlOneQuantity.Text )},,,,",
                $"Bowl 2,{bowlTwoComboBox.Text},,{( string.IsNullOrEmpty(bowlTwoQuantity.Text) ? string.Empty : bowlTwoQuantity.Text )},,,,",
                $"Bowl 3,{bowlThreeComboBox.Text},,{( string.IsNullOrEmpty(bowlThreeQuantity.Text) ? string.Empty : bowlThreeQuantity.Text )},,,,",
                $"Bowl 4,{bowlFourComboBox.Text},,{( string.IsNullOrEmpty(bowlFourQuantity.Text) ? string.Empty : bowlFourQuantity.Text )},,,,",
                $"Bowl 5,{bowlFiveComboBox.Text},,{( string.IsNullOrEmpty(bowlFiveQuantity.Text) ? string.Empty : bowlFiveQuantity.Text )},,,,",
                $"Bowl 6,{bowlSixComboBox.Text},,{( string.IsNullOrEmpty(bowlSixQuantity.Text) ? string.Empty : bowlSixQuantity.Text )},,,,",
                $"Bowl 7,{bowlSevenComboBox.Text},,{( string.IsNullOrEmpty(bowlSevenQuantity.Text) ? string.Empty : bowlSevenQuantity.Text )},,,,",
                $"Bowl 8,{bowlEightComboBox.Text},,{( string.IsNullOrEmpty(bowlEightQuantity.Text) ? string.Empty : bowlEightQuantity.Text )},,,,",
                $"Flange_Length,{FlangeOverhang.Text},,,,,,",
                $"Top_PEM_Safe_Z,{Top_PEM_Safety.Text},,,,,,",
                $"Bottom_PEM_Safe_Z,{Bottom_PEM_Saftey.Text},,,,,,",
                $"Flange_Unload_Direction_Up,{FlangePalletDirectionBox.Text.Last()},,,,,,",
                $"Pallet_Unload_Rotation,{Pallet_Outfeed_Rotation.Text},,,,,,",
                $"PEM,X,Y,Z,Bowl Number,,,"
            };

            foreach( var pem in m_PEMList )
            {
                var bowlNumber = "";
                foreach( Control control in Controls )
                {
                    if( control is GroupBox )
                    {
                        foreach( Control groupControl in control.Controls )
                        {
                            if( groupControl is ComboBox )
                            {
                                if( ( (ComboBox)groupControl ).SelectedIndex != -1 )
                                {
                                    if( ( (ComboBox)groupControl ).Text == pem.Name && m_SelectedQuantities[ groupControl.Tag.ToString( ) ] > 0 )
                                    {
                                        bowlNumber = ( (ComboBox)groupControl ).Tag.ToString( );
                                        m_SelectedQuantities[ groupControl.Tag.ToString( ) ] -= 1;
                                        pem.Bowl = int.Parse(bowlNumber);
                                        Console.WriteLine(int.Parse(bowlNumber) + " "+bowlNumber);
                                        fileOutputRows.Add( $"{pem.Name},,,,{bowlNumber},,," );
                                        fileOutputRows.Add( $",{string.Join( ",", pem.Coordinates )},,,," );
                                        

                                        break;
                                    }
                                }
                            }
                        }
                    }
                }
            }
            if (!uploading)
            {
            var dataAsBytes = fileOutputRows.SelectMany( s => Encoding.UTF8.GetBytes( s + "\r\n" ) ).ToArray( );

            Stream myStream;
          
                saveFileDialog.Filter = "CSV file (*.csv)|*.csv| All Files (*.*)|*.*"; ;
                saveFileDialog.RestoreDirectory = true;
                saveFileDialog.FileName = Path.GetFileNameWithoutExtension(m_FileLocation);
                try
                {
                    if( saveFileDialog.ShowDialog( ) == DialogResult.OK )
                    {
                        if( ( myStream = saveFileDialog.OpenFile( ) ) != null )
                        {
                            foreach( var data in dataAsBytes )
                            {
                                myStream.WriteByte( data );
                            }

                            myStream.Close( );
                        }
                    }
                }
                catch( Exception exception )
                {
                    MessageBox.Show( exception.ToString( ), "Uh oh!" );
                }
            }
        }

        private void Form1_DragDrop( object sender, DragEventArgs e )
        {
            var file = (string[ ])e.Data.GetData( DataFormats.FileDrop, false );

            FileValidation( file[ 0 ] );
        }

        private void Form1_DragEnter( object sender, DragEventArgs e )
        {
            e.Effect = e.Data.GetDataPresent( DataFormats.FileDrop ) ? DragDropEffects.Copy : DragDropEffects.None;
        }

        private void LoadFileButton_Click( object sender, EventArgs e )
        {
            openFileDialog.ShowDialog( );

            var file = openFileDialog.FileName.ToString( );

            FileValidation( file );
        }

        private void ProcessButton_Click( object sender, EventArgs e )
        {
            var valid = ValidFields( );

            if( valid && m_Revalidate )
            {
                processButton.Text = "Process";
                outputRichTextBox.Text = m_Output;
                m_Revalidate = false;
            }
            else if( valid  )
            {
                CreateOutputFile( );
            }
            else
            {
                processButton.Text = "Revalidate";
                m_Revalidate = true;
            }
        }

        private void BowlQuantities_TextChanged( object sender, EventArgs e )
        {
            var controls = GetControls( ControlSections.Robots );

            foreach( var control in controls )
            {
                if( control is ComboBox )
                {

                }
            }
            Console.WriteLine( ( (TextBox)sender ).Tag );
        }

        private void DatabaseSend(object sender, EventArgs e)
        {
            if (ValidFields() && !m_Revalidate) // Validate before sending
            {
                var cartTypeRegexPattern = new Regex(@"(\w{1}\d{2}\w{1}\d{2})(\w{1})(\d{1,2})");
                var infeedCart = cartTypeRegexPattern.Match(infeedCartComboBox.Text);
                var outfeedCart = cartTypeRegexPattern.Match(outfeedCartComboBox.Text);
                CreateOutputFile();

                DialogResult dialogResult = MessageBox.Show("Are you sure you want to send this PEM data to the production database?", "Warning", MessageBoxButtons.YesNo);
                if (dialogResult == DialogResult.Yes)
                {
                    try
                    {
                        SqlCommand sqlCommand = new SqlCommand("sp_UPSERT_PEM", conn);
                        sqlCommand.CommandType = System.Data.CommandType.StoredProcedure;
                        Console.WriteLine("UPSERT ATTEMPT");
                        sqlCommand.Parameters.Add(new SqlParameter("@BARCODE_NUMBER", barcodeNumberTextBox.Text));
                        sqlCommand.Parameters.Add(new SqlParameter("@CELL_1_HAEGER_PROGRAM", cellHaegerProgramOneTextBox.Text));
                        sqlCommand.Parameters.Add(new SqlParameter("@CELL_2_HAEGER_PROGRAM", cellHaegerProgramTwoTextBox.Text));
                        sqlCommand.Parameters.Add(new SqlParameter("@X_DIMENSION", xDimensionTextBox.Text));
                        sqlCommand.Parameters.Add(new SqlParameter("@Y_DIMENSION", yDimensionTextBox.Text));
                        sqlCommand.Parameters.Add(new SqlParameter("@Z_DIMENSION", zDimensionTextBox.Text));
                        sqlCommand.Parameters.Add(new SqlParameter("@FLANGE_DIRECTION", flangeDirectionComboBox.Text.Last()));
                        sqlCommand.Parameters.Add(new SqlParameter("@LENGTH_OPPOSITE_FLANGE", pemLengthOppositeFlangeTextBox.Text));
                        sqlCommand.Parameters.Add(new SqlParameter("@INFEED_CART_ORIENTATION", (infeedCartOrientation.Text == "P" ? "Portrait" : "Landscape")));
                        sqlCommand.Parameters.Add(new SqlParameter("@OUTFEED_CART_ORIENTATION", (outfeedCartOrientation.Text == "P" ? "Portrait" : "Landscape")));
                        sqlCommand.Parameters.Add(new SqlParameter("@INFEED_CART_SLOTS", infeedCartSlots.Text));
                        sqlCommand.Parameters.Add(new SqlParameter("@OUTFEED_CART_SLOTS", outfeedCartSlots.Text));
                        sqlCommand.Parameters.Add(new SqlParameter("@INFEED_CART_TYPE", infeedCartComboBox.Text));
                        sqlCommand.Parameters.Add(new SqlParameter("@OUTFEED_CART_TYPE", outfeedCartComboBox.Text));
                        sqlCommand.Parameters.Add(new SqlParameter("@X_EDGE_POS_CHECK1", xEdgeCheckPosition1TextBox.Text));
                        sqlCommand.Parameters.Add(new SqlParameter("@X_EDGE_POS_CHECK2", xEdgeCheckPosition2TextBox.Text));
                        sqlCommand.Parameters.Add(new SqlParameter("@Y_EDGE_POS_CHECK", yEdgeCheckPosition.Text));
                        sqlCommand.Parameters.Add(new SqlParameter("@TOOL_SIZE", toolSizeComboBox.Text.Last()));
                        sqlCommand.Parameters.Add(new SqlParameter("@RECIPE_TYPE", recipeTypeComboBox.Text.Last()));
                        sqlCommand.Parameters.Add(new SqlParameter("@EDGE_TYPE", edgeTypeComboBox.Text.Last()));
                        sqlCommand.Parameters.Add(new SqlParameter("@BOWL_1", bowlOneComboBox.Text));
                        sqlCommand.Parameters.Add(new SqlParameter("@BOWL_2", bowlTwoComboBox.Text));
                        sqlCommand.Parameters.Add(new SqlParameter("@BOWL_3", bowlThreeComboBox.Text));
                        sqlCommand.Parameters.Add(new SqlParameter("@BOWL_4", bowlFourComboBox.Text));
                        sqlCommand.Parameters.Add(new SqlParameter("@BOWL_5", bowlFiveComboBox.Text));
                        sqlCommand.Parameters.Add(new SqlParameter("@BOWL_6", bowlSixComboBox.Text));
                        sqlCommand.Parameters.Add(new SqlParameter("@BOWL_7", bowlSevenComboBox.Text));
                        sqlCommand.Parameters.Add(new SqlParameter("@BOWL_8", bowlEightComboBox.Text));
                        sqlCommand.Parameters.Add(new SqlParameter("@FLANGE_OVERHANG_LENGTH", FlangeOverhang.Text));
                        sqlCommand.Parameters.Add(new SqlParameter("@TOP_PEM_SAFE_Z", Top_PEM_Safety.Text));
                        sqlCommand.Parameters.Add(new SqlParameter("@BOTTOM_PEM_SAFE_Z",Bottom_PEM_Saftey.Text));
                        sqlCommand.Parameters.Add(new SqlParameter("@PALLET_UNLOAD_FLANGE_UP", FlangePalletDirectionBox.Text.Last()));
                        sqlCommand.Parameters.Add(new SqlParameter("@PALLET_OUTFEED_ROTATION", Pallet_Outfeed_Rotation.Text));
                        sqlCommand.Parameters.Add(new SqlParameter("@FLAG", true));

                        Console.WriteLine("EXECUTE SP");
                        sqlCommand.ExecuteNonQuery();
                    }
                    catch (Exception)
                    {
                        conn.Close();
                        throw;
                    }

                    try
                    { 
                    
                    // The code below inserts all the PEM positions in a seperate table with the table named after the part # 
                    Console.WriteLine("PEM LIST INSERT");
                        PEM badPem = null;
                        foreach (var pem in m_PEMList)
                        {
                          
                            if (skipFirst) // The first result has a bad PEM
                            {
                                if (pem.Name.Length > 16 )
                                {
                                    Console.Write(pem.Name);
                                }
                                else 
                                {
                                    skipFirst = false;
                                    var pemCommand = new SqlCommand("INSERT INTO \"" + barcodeNumberTextBox.Text + "\" VALUES(@PEM, @X, @Y,@Z,@Bowl)", conn);
                                    pemCommand.Parameters.AddWithValue("@PEM", pem.Name);
                                    pemCommand.Parameters.AddWithValue("@X", pem.Coordinates[0]); // X 
                                    pemCommand.Parameters.AddWithValue("@Y", pem.Coordinates[1]); // Y
                                    pemCommand.Parameters.AddWithValue("@Z", pem.Coordinates[2]); // Z
                                    pemCommand.Parameters.AddWithValue("@Bowl", pem.Bowl); // Bowl Number
                                    pemCommand.ExecuteNonQuery();
                                }
                            }
                            else
                            {
                                var pemCommand = new SqlCommand("INSERT INTO \"" + barcodeNumberTextBox.Text + "\" VALUES(@PEM, @X, @Y,@Z,@Bowl)", conn);
                                pemCommand.Parameters.AddWithValue("@PEM", pem.Name);
                                pemCommand.Parameters.AddWithValue("@X", pem.Coordinates[0]); // X 
                                pemCommand.Parameters.AddWithValue("@Y", pem.Coordinates[1]); // Y
                                pemCommand.Parameters.AddWithValue("@Z", pem.Coordinates[2]); // Z
                                pemCommand.Parameters.AddWithValue("@Bowl", pem.Bowl); // Bowl Number
                                pemCommand.ExecuteNonQuery();
                            }
                        }
                    }
                    catch (Exception)
                    {
                        conn.Close();
                        throw;
                    }
                    // Update recipe parameters
                    // Update/Replace PEM table
                    MessageBox.Show("RECIPE UPLOADED");
                }
                else if (dialogResult == DialogResult.No)
                {
                    conn.Close();
                    this.button1.Enabled = false;
                    uploading = false;
                }
            }
            else 
            {
                MessageBox.Show("Please Validate and Process Recipe Before Upload");
            }
        }

        private void DBconnect_Click(object sender, EventArgs e)
        {
            connectionstring = "";
            conn = new SqlConnection(connectionstring);
            try
            {
                conn.Open();
                MessageBox.Show("Connection Successful!");
                this.button2.Enabled = true;
                uploading = true;
            }
            catch (SqlException)
            {
                conn.Close();
                MessageBox.Show("Connection Failed, Check Database connection");
                uploading = false;

            }
        }

        private void Form1_Closing(object sender, System.ComponentModel.CancelEventArgs e)
        {
            conn.Close();
        }

        private void checkBox1_CheckedChanged_1(object sender, EventArgs e)
        {
            skipFirst = checkBox1.Checked;
        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {
            connectionstring = textBox1.Text;
            if (textBox1.Text == null || textBox1.Text.Length == 0)   // Check for empty field 
            {
                textBox1.Text = "DEFAULT";

                connectionstring = "";

            }
            else
            {
                connectionstring = textBox1.Text;
            }
        }

        private void Orientation_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void comboBox3_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }

        private void Discconect_Click(object sender, EventArgs e)
        {
            conn.Close();
            this.button2.Enabled = false;
            uploading = false;
        }

        private void c_TextChanged(object sender, EventArgs e)
        {

        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void flange_Overhang(object sender, EventArgs e)
        {

        }

        private void label34_Click(object sender, EventArgs e)
        {

        }

        private void label35_Click(object sender, EventArgs e)
        {

        }

        private void label36_Click(object sender, EventArgs e)
        {

        }

        private void checkBox2_CheckedChanged(object sender, EventArgs e)
        {
         
        }

        private void label27_Click(object sender, EventArgs e)
        {

        }

        private void label41_Click(object sender, EventArgs e)
        {

        }

        private void label42_Click(object sender, EventArgs e)
        {

        }

        private void label41_Click_1(object sender, EventArgs e)
        {

        }

        private void pemLengthOppositeFlangeTextBox_TextChanged(object sender, EventArgs e)
        {

        }

        private void Pallet_Unload_Rotation_TextChanged(object sender, EventArgs e)
        {

        }
    }
}
