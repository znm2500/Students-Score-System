   M S B A M L   `   `   `  ����B  >WpfApp1, Version=1.0.0.0, Culture=neutral, PublicKeyToken=null!clr-namespace:WpfApp1WpfApp1  5   	       WpfApp1.MainWindow   5   	   W SPresentationCore, Version=8.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35\ XPresentationFramework, Version=8.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35e aSystem.Windows.Controls.Ribbon, Version=8.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089R NWindowsBase, Version=8.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35F 9http://schemas.microsoft.com/winfx/2006/xaml/presentation     5   	   R NSystem.Xaml, Version=8.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089:x,http://schemas.microsoft.com/winfx/2006/xaml     5   	   @d2http://schemas.microsoft.com/expression/blend/2008     5   	   Jmc;http://schemas.openxmlformats.org/markup-compatibility/2006     5   	   'localclr-namespace:WpfApp1     5   	     � Title$  	主界面��5   	   $	��570��6   $	��960��6"   .��5
      �  � RowDefinitions 5   
   �� 5      $��30 �6   6*   �� 5      $	��400 �6   6+   .+�5   
   �� -   5        $��学生更改��5      $��80��6!   $
��Left=�6A   $��	160,0,0,0q�6\   $��30��6o   $	��Top=�6{   �� 5   
   -   5        $��分数更改��5      $��80��6    $��30��6+   $��Center=�67   $
��Left=�6R   $��80,0,0,0q�6�   �� 5   
   -   5        $��分数刷新��5      $��80��6    $��30��6+   $��Center=�67   $
��Left=�6R   6m   &   System.Windows.Controls.DataGrid  5   
   -   5          Name$ ranking��5   m   $��1��6   $
��Left=�6!   $	��950��6<   $��
0,0,0,-113q�6H   $��2��6\      AutoGenerateColumns$ False��6�   ��6   ��    Columns 5      0  *System.Windows.Controls.DataGridTextColumn  5      ,  &System.Windows.Controls.DataGridColumn   Header$	 组��6%      Width5  /System.Windows.Controls.DataGridLengthConverter$ 30 6J      
IsReadOnly$
 True��6U   1  +System.Windows.Controls.DataGridBoundColumn   Binding 60   �� *group+6U     5      $ 名字��6%   $ 60 6K   $
 True��6V    62   �� *name+6V     5      $ 分数��6%   $ 90 6L   $
 True��6W    62   �� *score+6W     5      $ 加分记录��6%   $	 750 6O   $
 True��6[    64   �� *record+6[   5      �� 5   
   -   5        $��小组排名��5      $��80��6!   $��30��6,   $��Center=�69   $
��Left=�6T   $��	240,0,0,0q�6o   6�   �� 5   
   -   5        $��导出表格��5      $��80��6!   $��30��6,   $��Center=�69   $
��Left=�6T   $��	320,0,0,0q�6o   6�   5"      5'      