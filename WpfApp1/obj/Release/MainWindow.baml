   M S B A M L   `   `   `  ���� O  KExcelGenerator2Swift, Version=1.0.0.0, Culture=neutral, PublicKeyToken=nullA%clr-namespace:ExcelGenerator2SwiftAppExcelGenerator2SwiftApp  (    "ExcelGenerator2SwiftApp.MainWindow   R NWindowsBase, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35W SPresentationCore, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35\ XPresentationFramework, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35D 9http://schemas.microsoft.com/winfx/2006/xaml/presentation    R NSystem.Xaml, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089:x,http://schemas.microsoft.com/winfx/2006/xaml     @d2http://schemas.microsoft.com/expression/blend/2008     Jmc;http://schemas.openxmlformats.org/markup-compatibility/2006     7local%clr-namespace:ExcelGenerator2SwiftApp       � Title$"  ExcelGenerator2Swift - v1.00��$
��1227��$
��1194�� � Icon$ 
favico.ico��
���   �.��� $��
0,0,114,98q� � RowDefinitions �� �� $��0* �.+�� $��0,0,-114,-99q�$��2�� �� $��35 ��� $
��114* ��� $
��17.5 ��� $��35 ��� $��15 ��� $��35 ��� $��15 ��� $��35 ��� $��15 ��� $��35 ��� $��15 ��� $��35 ��� $��15 ��� $��35 ��� $��35 ��� $	��14* ��� $	��96* ��� $��15 ��� $	��91* ��� $��15 ��� $��15 ��� $��35 ��� $��15 ��� $��35 ��� $��35 ��� $	��92* ��� $��15 ��� $	��94* ��� $��15 ��� $��35 ��� $��35 ��� $��35 ��� $	��31* � � ColumnDefinitions �� $��6* ��� $	��70* ��� $
��100* ��� $	��31* ��� $
��100* ��� $	��32* ��� $	��99* ��� $	��29* ��� $
��100* ��� $	��30* ��� $
��100* ��� $	��21* ��� $
��118* ��� $	��26* ��� $	��88* ��� $	��96* ��� $��6* ��� $
��141* ��� $	��11* �.+��� -   $��CompanyName��$-��'Informatique CDC - Projet SIGMA - Swift��$��8��$��24(�$
��Bold"�
���  ��$��8��$
��Wrap=� �� TextAlignment$ Center=�$��Italic$�$��1����Z� �� -   $��About��$��2��$
��Left=�$��1��$	��Top=�$��41��$	��193��$��3��.���� $��Images/icdc.png���� $��	Suffix : ��$��3��$��14��
���  ��$��	Underline{�$
��Bold"�$	��0,5q�$��6���� -    ��Name$ txbSuffix_0��$��3��$��15�� �� TextWrapping$
 Wrap=�$��A��$
��Bold"��� $��Path : ��$��3��$��1��
���  ��$��	Underline{�$
��Bold"��� -   $��txbPathSwift_0��$��3��$��2��$
 Wrap=�$*��$C:\Developpement\Tests_dev\Swift\in\��$��11��$
��Bold"� � ToolTip$! Suffix added on Swift files���� -   $��btnBrowse_0��$��3��$	��...��$��17�����	#FFE2E220$
��Bold"�$1 +Browser to select the folder of Swift files���� -   $��btnDuplicate_0��$��5��$��Duplicate files��$��17�����	#FFE2E220$
��Bold"�$' !Duplicate Swift files with suffix���� $��Filter file : ��$��7��$��1��
���  ��$��	Underline{�$
��Bold"��� -   $��txbFilterFileSwift_0��$��7��$��2��$
 Wrap=�$*��$C:\Developpement\Tests_dev\Swift\in\��$��11��$
��Bold"�$! Suffix added on Swift files���� -   $��btnBrowseFilterFile_0��$��7��$	��...��$��14�����	#FFE2E220$
��Bold"�$( "Browser to select the filter files���� -	   $��btnFilter_0��$��7��$��Filter files��$��17�����	#FFE2E220$
��Bold"�$$ Filter Swift files with suffix���� $��File : ��$��9��$��1��
���  ��$��	Underline{�$
��Bold"��� -
   $��txbPathExcelFile��$��2��$��9��$
 Wrap=�$/��)C:\Developpement\Tests_dev\TestSwift.xlsx��$��14��$
��Bold"�&   System.Windows.Controls.DataGrid  -   $��
dataGrid_1��$��2��$��11��$��14��$��6���� -   $��btnBrowse_1��$	��...��$��17��$��9�����	#FFE2E220$
��Bold"�$&  Browser to select the Excel file���� -   $��	btnRead_1��$��Read Excel file��$��17��$��11�����	#FF77E67C$
��Bold"�$ Read the Excel file���� -   $��btnGenerate_1��$��Generate Swift files��$��17��$��13�����	#FF66B4EA$
��Bold"�$. (Write text file according the Excel file���� $��Path : ��$��18��$��1��
���  ��$��	Underline{�$
��Bold"�$��0,10,0,0q�$��2���� -   $ txbPathDirectory��$��2��$��18��$
 Wrap=�$!��C:\Developpement\Tests_dev\��$��11��$
��Bold"�$��0,3,0,0q��� -    ��Name$ btnBrowse_2��$	��...��$��16��$��18�����	#FFE2E220$
��Bold"�$2 ,Browser to select the folder with text files��$��2��$��5,0,2,3q�$��32��$��Bottom=��� -   $ 	btnRead_2��$��Read Text files��$��17��$��21�����	#FF77E67C$
��Bold"�$% Read the folder with text files���� -   $ btnGenerate_2��$��Generate Excel file��$��17��$��23�����	#FF66B4EA$
��Bold"�$; 5Write the Excel file 'ExportExcel.xlsx' in the folder��$��0,0,1,0q�  -   	  Name$	 
dataGrid_2��$��2��$��21��$��14��$��5���� $��Extension : ��$��18��$��14��
���  ��$��	Underline{�$
��Bold"�$	��0,5q�$��2���� -   $ txbExtension��$��15��$��18��$
 Wrap=�$
��.txt��$
��Bold"��� $��Path : ��$��27��$��1��
���  ��$��	Underline{�$
��Bold"��� -   $��txbPathSwift��$��2��$��27��$
 Wrap=�$!��C:\Developpement\Tests_dev\��$��11��$
��Bold"��� -   $��btnBrowse_3��$	��...��$��17��$��27�����	#FFE2E220$
��Bold"�$0 *Browser to select the folder of test files���� -   $ btnMT502��$��MT502��$��2��$��29�����	#FF66B4EA$
��Bold"�$* $Write MT502 Test files in the folder���� -   $ btnMT54X��$��MT54X��$��4��$��29�����	#FF66B4EA$
��Bold"�$* $Write MT54X Test files in the folder���� -   $ btnMT598��$��MT598��$��6��$��29�����	#FF66B4EA$
��Bold"�$* $Write MT598 Test files in the folder���� -   $ btnMT304��$��MT304��$��8��$��29�����	#FF66B4EA$
��Bold"�$* $Write MT304 Test files in the folder���� -   $ btnUBIX��$
��UBIX��$��10��$��29�����	#FF66B4EA$
��Bold"�$) #Write UBIX Test files in the folder���� -   $ btnGenerate_3��$��Generate 1000 Test files��$��17��$��29�����	#FF66B4EA$
��Bold"�$- 'Write the 1000 Test files in the folder���� $��Nbre : ��$��27��$��14��
���  ��$��	Underline{�$
��Bold"�$��2���� -   $ txbNbreTestFiles��$��15��$��27��$
 Wrap=�$
��1000��$
��Bold"��� -   $��btnClose��$��Close��$��17��$��31�����	#FFDE24CD$
��Bold"�$ Close the application��