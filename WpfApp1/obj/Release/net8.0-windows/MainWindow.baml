   M S B A M L   `   `   `  ���� B  >WpfApp1, Version=1.0.0.0, Culture=neutral, PublicKeyToken=null!clr-namespace:WpfApp1WpfApp1      WpfApp1.MainWindow   W SPresentationCore, Version=8.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35\ XPresentationFramework, Version=8.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35e aSystem.Windows.Controls.Ribbon, Version=8.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089R NWindowsBase, Version=8.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35F 9http://schemas.microsoft.com/winfx/2006/xaml/presentation     R NSystem.Xaml, Version=8.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089:x,http://schemas.microsoft.com/winfx/2006/xaml     @d2http://schemas.microsoft.com/expression/blend/2008     Jmc;http://schemas.openxmlformats.org/markup-compatibility/2006     'localclr-namespace:WpfApp1       � Title$  	主界面��$	��570��$	��960��.���  � RowDefinitions �� $��30 ��� $	��400 �.+��� -   $��学生更改��$��80��$
��Left=�$��	160,0,0,0q�$��30��$	��Top=��� -   $��分数更改��$��80��$��30��$��Center=�$
��Left=�$��80,0,0,0q��� -   $��分数刷新��$��80��$��30��$��Center=�$
��Left=�&   System.Windows.Controls.DataGrid  -     Name$ ranking��$��1��$
��Left=�$	��950��$��
0,0,0,-113q�$��2��   AutoGenerateColumns$ False������    Columns 0  *System.Windows.Controls.DataGridTextColumn  ,  &System.Windows.Controls.DataGridColumn   Header$	 组��   Width5  /System.Windows.Controls.DataGridLengthConverter$ 30    
IsReadOnly$
 True��1  +System.Windows.Controls.DataGridBoundColumn   Binding �� *group+  $ 名字��$ 60 $
 True�� �� *name+  $ 分数��$ 90 $
 True�� �� *score+  $ 加分记录��$	 750 $
 True�� �� *record+�� -   $��小组排名��$��80��$��30��$��Center=�$
��Left=�$��	240,0,0,0q��� -   $��导出表格��$��80��$��30��$��Center=�$
��Left=�$��	320,0,0,0q�