   M S B A M L   `   `   `  ���� G  CLAPreparator, Version=1.0.0.0, Culture=neutral, PublicKeyToken=null+clr-namespace:LAPreparatorLAPreparator  E'clr-namespace:LAPreparator.UIComponentsLAPreparator.UIComponents  O Kmscorlib, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e0891&clr-namespace:System;assembly=mscorlibSystem     LAPreparator.MainWindow   W SPresentationCore, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35\ XPresentationFramework, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35e aSystem.Windows.Controls.Ribbon, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089R NWindowsBase, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35F 9http://schemas.microsoft.com/winfx/2006/xaml/presentation     R NSystem.Xaml, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089:x,http://schemas.microsoft.com/winfx/2006/xaml     @d2http://schemas.microsoft.com/expression/blend/2008     Jmc;http://schemas.openxmlformats.org/markup-compatibility/2006     ,localclr-namespace:LAPreparator     _ [MaterialDesignThemes.Wpf, Version=4.3.0.3, Culture=neutral, PublicKeyToken=df2a72020bd7962aFmaterialDesign1http://materialdesigninxaml.net/winfx/xaml/themes  :uicomponents'clr-namespace:LAPreparator.UIComponents  3System&clr-namespace:System;assembly=mscorlib    � Title$  
MainWindow��$	��450��$	��800�� � Icon$ /resources/logoapmm.jpg�� � 	Resources    logo �� 	UriSource  	titleLogo  ButtonHeight  SideBarWidth   BooleanToVisibilityConverter�� %�   &	        &	 %     &	 H     &	 T     &	 a     �� $ /Resources/maersklogo.png?��� $ /Resources/logoapmm.jpg?�X�30Q�X�250Q��� .��� .+�]� $��Stretch=�$
��Left=�$	��250��   MaterialDesignDarkBackground#���  .+��� $	��Top=�#��[  $
��Left=�$	��Top=�$��30��� .+��� #���  $	��Top=�$	��150��$	��120�� �� Orientation$ Vertical=�#��[ $
��Left=���u� 
 u� Top$ 30Q� u� Left$ 0Q�.+��� -   $��ConfigurationMenuButton��  AttentionToActionBrush#���  $��30��$
��Left=�.���� $ 
Horizontal=�.+�'  !MaterialDesignThemes.Wpf.PackIcon     Kind$ SettingsOutline=��� $��Configuration��$	��Top=�   MaterialDesignDarkForeground#���  ��u� $ 10Q�$ -2Q��� -    ��Name$ ConfigurationMenuStack��$y�	Collapsed=�.+��� -   $��MessageBodyButton��
�������#��[ $
��Left=�.���� $	��220��$��LeftToRight=�$ 
Horizontal=�.+�  $ MessageBadgeOutline=���u� $ 20Q��� $��
LA Message��$	��Top=�$	��-10q�#���  ��u� $ 10Q�$ -2Q��� 
�������#��[ $
��Left=�.���� $	��220��$��LeftToRight=�$ 
Horizontal=�.+�  $ MessageBadgeOutline=���u� $ 20Q��� $��Sail Confirmaton Message��$	��Top=�$	��-10q�#���  )  #MaterialDesignThemes.Wpf.HintAssist	   
HelperText$7	 1Message Template for sending Sailing Confirmation��$
��Left=���u� $ 10Q�$ -2Q�
 u� Right$	
 -20Q��� -   $��AddressBookButton��
�������#��[ $
��Left=�.���� $	��200��$��LeftToRight=�$ 
Horizontal=�.+�  $ ContactMailOutline=���u� $ 20Q��� $��Address Book��$	��Top=�$	��-10q�#���  ��u� $ 10Q�$ -2Q���  , (MaterialDesignDataGridRowHoverBackground#���  $	��Top=�$	��150��$	��120��$ Vertical=�#��[ $
��Left=���u� $	 200Q�.+��� -   $��EmailMenuButton��#���  $��30��$
��Left=�.���� $ 
Horizontal=�.+�  $ EmailMultipleOutline=��� $��Emails��$	��Top=�#���  ��u� $ 10Q�$ -2Q��� -   $ EmailMenuStack��$y�	Collapsed=�.+��� -   $��SendLaButton��
�������#��[ $
��Left=�.���� $	��220��$��LeftToRight=�$ 
Horizontal=�.+�  $ SendOutline=���u� $ 20Q��� $��Send LA��$	��Top=�$	��-10q�#���  ��u� $ 10Q�$ -2Q��� 
�������#��[ $
��Left=�.���� $	��220��$��LeftToRight=�$ 
Horizontal=�.+�  $ SendLock=���u� $ 20Q��� $��Send Sailing Confirmation��$	��Top=�$	��-10q�#���     Hint$7 1Message Template for sending Sailing Confirmation����u� $ 10Q�$ -2Q�$	
 -10Q�$	��220��]� -   $��	BodyPanel��$��Stretch=�$��Stretch=�
���������u� $	 250Q�