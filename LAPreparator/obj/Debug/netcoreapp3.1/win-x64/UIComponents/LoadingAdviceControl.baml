   M S B A M L   `   `   `  ����G  CLAPreparator, Version=1.0.0.0, Culture=neutral, PublicKeyToken=nullE'clr-namespace:LAPreparator.UIComponentsLAPreparator.UIComponents  5      O Kmscorlib, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e0891&clr-namespace:System;assembly=mscorlibSystem 5      4    .LAPreparator.UIComponents.LoadingAdviceControl   5      W SPresentationCore, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35\ XPresentationFramework, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35e aSystem.Windows.Controls.Ribbon, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089R NWindowsBase, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35F 9http://schemas.microsoft.com/winfx/2006/xaml/presentation     5      R NSystem.Xaml, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089:x,http://schemas.microsoft.com/winfx/2006/xaml     5      Jmc;http://schemas.openxmlformats.org/markup-compatibility/2006     5      @d2http://schemas.microsoft.com/expression/blend/2008     5      9local'clr-namespace:LAPreparator.UIComponents     5      _ [MaterialDesignThemes.Wpf, Version=4.3.0.3, Culture=neutral, PublicKeyToken=df2a72020bd7962aFmaterialDesign1http://materialdesigninxaml.net/winfx/xaml/themes  5      3System&clr-namespace:System;assembly=mscorlib  5        � 	Resources  5         	TopMargin  LeftLabelMargin  LeftTextBoxMargin��%m   &	        &	      &	 4     X�5   
   10Q�6.   X�5   
   50Q�64   X�5   
   100Q�67   5      .��5      �    MaterialDesignCardBackground#���  6   .+�5   
   #  MaterialDesignThemes.Wpf.Card  $	��Top=�6   $��Stretch=�67   $��50��6U   ��5      u� 5       u� Left$ 0Q�6   
 u� Top$ 10Q�6%   60   .��5      �� $��LA Email Panel��6   $��Center=�6-   $��Center=�6J     MaterialDesignShadowBrush#���  6e   6�   5        5   
   $��Stretch=�6   
���������5      u� 5      $ 70Q�6   6'   .��5      � $��Stretch=�6   .+�5      � $��Center=�6   .+�5      �� $
��File��6    ! MaterialDesignLightForeground#���  6'   $
��Left=�6d   $	��Top=�6   ��5      u� 5      # [  6   $ 50Q�6=   6I   5!      �� 5"      -   5         ��Name$ FileTextBox��5"      $
��Left=�6/   $	��Top=�6J   $	��200��6b   .��6�   ��5$      u� 5%      # [  6   # [ 6=   6i   .��5&      5'      �� 5(      -   5        $��BrowseButton��5(   �   $	��Top=�6   $
��Left=�61   $	��100��6L   $��Browse��6X     PrimaryHueDarkBrush#���  6i   ��5)      u� 5*      # [  6$   $	 320Q�6E   6R   5,      �� 5-      $��Sheet��6   $
��Left=�6(   $	��Top=�6C   ��5.      u� 5/      # [ 6$   $ 40Q�6L   6W   51      �� 52      -   5        $��SheetNameTextBox��52   g   $	��Top=�6   $
��Left=�62   $	��200��6M   $��Sheet1��6Y   .��6   ��53      u� 54      # [ 6$   $ 40Q�6N   6Y   .��55   &   56      �� 57      $��Range��6   $��30��6(   $	��Top=�64   $��40��6L   $
��Left=�6W   ��58      u� 59      # [ 6$   $ 70Q�6L   6W   5;      �� 5<      -   5        $��RangeTextBox��5<   g   $	��200��6   $	��Top=�6&   $
��Left=�6>   $��A15:AZ��6Y   .��6{   ��5=      u� 5>      # [ 6$   $ 70Q�6N   6Y   .��6t   5?      �� 5@      -   5        $��
LoadButton��5@   o   $	��200��6   $	��Top=�6%   $
��Left=�6=   
����ح�$
��Load��6�   ��5A      u� 5B      $	 100Q�6$   $	 100Q�6/   6;   5C      �� 5D      -   5        $��
SendButton��5D   s   $	��200��6   $	��Top=�6)   $
��Left=�6A   
����ح�$
��Send��6�   ��5E      u� 5F      $	 100Q�6(   $	 140Q�63   6?   5H      5I      � 5J      -   5         �Name$ GroupByPanel��5J      $��Stretch=�62   $	��Top=�6P   
�������$	��100��6   $��LeftToRight=�6�    � Orientation$ 
Horizontal=�6�   ��5K      u� 5L      $	 200Q�6$   60    � 
ItemHeight$ 30��5N   D    � 	ItemWidth$ 95��5O   B   5P      &   System.Windows.Controls.DataGrid  5Q      -   5          Name$ DataGrid��5Q      $��Stretch=�6-   $��Stretch=�6I   	   AutoGenerateColumns$	 False��6g      MaterialDesignChipBackground#���  6�   ��5R      u� 5S      $	 300Q�6$   60   5[      5\      5^      5_      5`      