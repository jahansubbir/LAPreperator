   M S B A M L   `   `   `  ���� G  CLAPreparator, Version=1.0.0.0, Culture=neutral, PublicKeyToken=nullE'clr-namespace:LAPreparator.UIComponentsLAPreparator.UIComponents  O Kmscorlib, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e0891&clr-namespace:System;assembly=mscorlibSystem 4    .LAPreparator.UIComponents.LoadingAdviceControl   W SPresentationCore, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35\ XPresentationFramework, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35e aSystem.Windows.Controls.Ribbon, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089R NWindowsBase, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35F 9http://schemas.microsoft.com/winfx/2006/xaml/presentation     R NSystem.Xaml, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089:x,http://schemas.microsoft.com/winfx/2006/xaml     Jmc;http://schemas.openxmlformats.org/markup-compatibility/2006     @d2http://schemas.microsoft.com/expression/blend/2008     9local'clr-namespace:LAPreparator.UIComponents     _ [MaterialDesignThemes.Wpf, Version=4.3.0.3, Culture=neutral, PublicKeyToken=df2a72020bd7962aFmaterialDesign1http://materialdesigninxaml.net/winfx/xaml/themes  3System&clr-namespace:System;assembly=mscorlib    � 	Resources     	TopMargin  LeftLabelMargin  LeftTextBoxMargin��%C   &	        &	      &	      X�10Q�X�50Q�X�100Q�.���    MaterialDesignCardBackground#���  .+�#  MaterialDesignThemes.Wpf.Card  $	��Top=�$��Stretch=�$��50����u�  u� Left$ 0Q�
 u� Top$ 10Q�.���� $��LA Email Panel��$��Center=�$��Center=�  MaterialDesignShadowBrush#���    $��Stretch=�
���������u� $ 70Q�.��� $��Stretch=�.+�� $��Center=�.+��� $
��File�� ! MaterialDesignLightForeground#���  $
��Left=�$	��Top=���u� # [  $ 50Q��� -    ��Name$ FileTextBox��$
��Left=�$	��Top=�$	��200��.����u� # [  # [ .���� -   $��BrowseButton��$	��Top=�$
��Left=�$	��100��$��Browse��  PrimaryHueDarkBrush#���  ��u� # [  $	 320Q��� $��Sheet��$
��Left=�$	��Top=���u� # [ $ 40Q��� -   $��SheetNameTextBox��$	��Top=�$
��Left=�$	��200��$��Sheet1��.����u� # [ $ 40Q�.���� $��Range��$��30��$	��Top=�$��40��$
��Left=���u� # [ $ 70Q��� -   $��RangeTextBox��$	��200��$	��Top=�$
��Left=�$��A15:AZ��.����u� # [ $ 70Q�.���� -   $��
LoadButton��$	��200��$	��Top=�$
��Left=�
����ح�$
��Load����u� $	 100Q�$	 100Q��� -   $��
SendButton��$	��200��$	��Top=�$
��Left=�
����ح�$
��Send����u� $	 100Q�$	 140Q�� -    �Name$ GroupByPanel��$��Stretch=�$	��Top=�
�������$	��100��$��LeftToRight=� � Orientation$ 
Horizontal=���u� $	 200Q� � 
ItemHeight$ 30�� � 	ItemWidth$ 95��&   System.Windows.Controls.DataGrid  -     Name$ DataGrid��$��Stretch=�$��Stretch=�	   AutoGenerateColumns$	 False��   MaterialDesignChipBackground#���  ��u� $	 300Q�