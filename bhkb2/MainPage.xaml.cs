using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices.WindowsRuntime;
using Windows.Foundation;
using Windows.Foundation.Collections;
using Windows.Storage;
using Windows.Storage.Pickers;
using Windows.UI.Xaml;
using Windows.UI.Xaml.Controls;
using Windows.UI.Xaml.Controls.Primitives;
using Windows.UI.Xaml.Data;
using Windows.UI.Xaml.Input;
using Windows.UI.Xaml.Media;
using Windows.UI.Xaml.Navigation;

// The Blank Page item template is documented at http://go.microsoft.com/fwlink/?LinkId=402352&clcid=0x409

namespace bhkb2
{
    /// <summary>
    /// An empty page that can be used on its own or navigated to within a Frame.
    /// </summary>
    public sealed partial class MainPage : Page
    {
        public MainPage()
        {
            this.InitializeComponent();
        }

        private async void button_Click(object sender, RoutedEventArgs e)
        {
            FileOpenPicker openFile = new FileOpenPicker();
            openFile.SuggestedStartLocation = PickerLocationId.Downloads;
            openFile.ViewMode = PickerViewMode.List;
            openFile.FileTypeFilter.Add(".xls");

            StorageFile file = await openFile.PickSingleFileAsync();
            if (file != null)
            {
                temp_message.Text = "你所选择的文件是： " + file.Name;

                byte b1, b2 = 0;
                int num_of_str = 0;
                byte[][] strs = new byte[100][];
                int[,] str_id = new int[100, 100];
                using (Stream file1 = await file.OpenStreamForReadAsync())
                {
                    using (BinaryReader read = new BinaryReader(file1))
                    {
                        //read.Read(data, 0, 10000);
                        try
                        {
                            while (true)
                            {
                                b1 = read.ReadByte();

                                if (b1 == 0 && b2 == 0xfc)
                                {
                                    //read strs
                                    read.ReadInt16();
                                    read.ReadInt32();
                                    num_of_str = read.ReadInt32();

                                    for (int i = 0; i < num_of_str; i++)
                                    {
                                        int length = read.ReadInt16();
                                        int long_char = read.ReadByte();
                                        length = length << long_char;
                                        strs[i] = read.ReadBytes(length);
                                    }

                                    b2 = 0;
                                    continue;
                                }
                                if (b1 == 0 && b2 == 0xfd)
                                {
                                    b2 = 0;
                                    read.ReadInt16();
                                    int row = read.ReadInt16();
                                    int col = read.ReadInt16();
                                    int s = read.ReadInt16();

                                    int id = read.ReadInt32();
                                    str_id[row, col] = id;

                                    continue;
                                }
                                b2 = b1;
                            }
                        }
                        catch (Exception exce)
                        {

                        }
                    }
                }
                InitRows(6, grid1);
                InitColumns(7, grid1);
                for (int i = 2; i < 8; i++)
                {
                    for (int j = 2; j < 9; j++)
                    {

                        TextBlock block = new TextBlock();

                        block.Text = System.Text.Encoding.Unicode.GetString(strs[str_id[i, j]]);

                        grid1.Children.Add(block);
                        Grid.SetRow(block, i - 2);
                        Grid.SetColumn(block, j - 2);
                        //temp_message.Text = ""+ str_id[i,j] + " "+ System.Text.Encoding.Unicode.GetString(strs[str_id[i, j]]);
                    }
                }

            }
            else
            {
                temp_message.Text = "打开文件操作被取消。";
            }
        }
        private void InitRows(int rowCount, Grid g)
        {
            while (rowCount-- > 0)
            {
                RowDefinition rd = new RowDefinition();
                rd.Height = new GridLength();
                g.RowDefinitions.Add(rd);
            }
        }
        private void InitColumns(int colCount, Grid g)
        {
            while (colCount-- > 0)
            {
                ColumnDefinition rd = new ColumnDefinition();
                rd.Width = new GridLength(100);
                g.ColumnDefinitions.Add(rd);
            }
        }
    }
}
