   M S B A M L   `   `   `  ����G  CLAPreparator, Version=1.0.0.0, Culture=neutral, PublicKeyToken=null+clr-namespace:LAPreparatorLAPreparator  5   	   E'clr-namespace:LAPreparator.UIComponentsLAPreparator.UIComponents  5   	   O Kmscorlib, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e0891&clr-namespace:System;assembly=mscorlibSystem 5	   	       LAPreparator.MainWindow   5   	   W SPresentationCore, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35\ XPresentationFramework, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35e aSystem.Windows.Controls.Ribbon, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089R NWindowsBase, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35F 9http://schemas.microsoft.com/winfx/2006/xaml/presentation     5   	   R NSystem.Xaml, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089:x,http://schemas.microsoft.com/winfx/2006/xaml     5   	   @d2http://schemas.microsoft.com/expression/blend/2008     5   	   Jmc;http://schemas.openxmlformats.org/markup-compatibility/2006     5   	   ,localclr-namespace:LAPreparator     5   	   _ [MaterialDesignThemes.Wpf, Version=4.3.0.3, Culture=neutral, PublicKeyToken=df2a72020bd7962aFmaterialDesign1http://materialdesigninxaml.net/winfx/xaml/themes  5   	   :uicomponents'clr-namespace:LAPreparator.UIComponents  5   	   3System&clr-namespace:System;assembly=mscorlib  5	   	     � Title$  
MainWindow��5   	   $	��450��6   $	��800��6)    � Icon$ /resources/logoapmm.jpg��65    � 	Resources 5         logo �� 	UriSource  	titleLogo  ButtonHeight  SideBarWidth   BooleanToVisibilityConverter�� 5   
   %�   &	        &	 3     &	 d     &	 ~     &	 �     �� 5      $ /Resources/maersklogo.png?�6'   �� 5      $ /Resources/logoapmm.jpg?�6,   X�5      30Q�65   X�5      250Q�66   �� 5      6+   5      .��5      � .+�5   
   ]� $��Stretch=�6   $
��Left=�60   $	��250��6K      MaterialDesignDarkBackground#���  6W   .+�5      �� $	��Top=�6   #��[  6,   $
��Left=�6K   $	��Top=�6f   $��30��6{   6�   � 5      .+�5      �� #���  6   $	��Top=�6U   $	��150��6n   $	��120��6~    �� Orientation$ Vertical=�6�   #��[ 6�   $
��Left=�6�   ��5      u� 5      
 u� Top$ 30Q�6     u� Left$ 0Q�6)   64   .+�5       �� -   5        $��ConfigurationMenuButton��5    P     AttentionToActionBrush#���  6   $��30��6p   $
��Left=�5!   `   .��5"      �� $ 
Horizontal=�6!   .+�5#      '  !MaterialDesignThemes.Wpf.PackIcon     Kind$ SettingsOutline=�62   6L   �� 5$      $��Configuration��6    $	��Top=�68      MaterialDesignDarkForeground#���  6Q   ��5%      u� 5&   "   $ 10Q�6,   $ -2Q�66   6A   5(      5*      5+      �� 5,      -   5         ��Name$ ConfigurationMenuStack��5,      $y�	Collapsed=�6=   .+�5-      �� -   5        $��MessageBodyButton��5-   �   
�������#��[ 68   $
��Left=�6_   .��5.      �� $	��220��6%   $��LeftToRight=�61   $ 
Horizontal=�6M   .+�5/        $ MessageBadgeOutline=�66   ��50   "   u� 51   &   $ 20Q�60   6<   53      �� 54      $��
LA Message��6$   $	��Top=�69   $	��-10q�6Q   #���  6_   ��55   "   u� 56   &   $ 10Q�60   $ -2Q�6:   6E   58      59      5:      �� 5;      
�������#��[ 67   $
��Left=�6^   .��5<      �� $	��220��6%   $��LeftToRight=�61   $ 
Horizontal=�6M   .+�5=        $ MessageBadgeOutline=�66   ��5>   "   u� 5?   &   $ 20Q�60   6<   5A      �� 5B      $��Sail Confirmaton Message��6$   $	��Top=�6G   $	��-10q�6_   #���  6n   )  #MaterialDesignThemes.Wpf.HintAssist	   
HelperText$7	 1Message Template for sending Sailing Confirmation��6�   $
��Left=�6  ��5C   "   u� 5D   &   $ 10Q�60   $ -2Q�6:   
 u� Right$	
 -20Q�6C   6Q   5F      5G      5H      �� 5I      -   5        $��AddressBookButton��5I   �   
�������#��[ 68   $
��Left=�6_   .��5J      �� $	��200��6%   $��LeftToRight=�61   $ 
Horizontal=�6M   .+�5K        $ ContactMailOutline=�66   ��5L   "   u� 5M   &   $ 20Q�60   6<   5O      �� 5P      $��Address Book��6$   $	��Top=�6;   $	��-10q�6S   #���  6a   ��5Q   "   u� 5R   &   $ 10Q�60   $ -2Q�6:   6E   5T      5U      5V      5W      5Z      �� 5[       , (MaterialDesignDataGridRowHoverBackground#���  6   $	��Top=�6a   $	��150��6z   $	��120��6�   $ Vertical=�6�   #��[ 6�   $
��Left=�6�   ��5\      u� 5]      $	 200Q�6    6,   .+�5_      �� -   5        $��EmailMenuButton��5_   P   #���  6   $��30��6h   $
��Left=�5`   `   .��5a      �� $ 
Horizontal=�6!   .+�5b        $ EmailMultipleOutline=�62   6Q   �� 5c      $��Emails��6    $	��Top=�61   #���  6J   ��5d      u� 5e   "   $ 10Q�6,   $ -2Q�66   6A   5g      5i      5j      �� 5k      -   5        $ EmailMenuStack��5k      $y�	Collapsed=�65   .+�5l      �� -   5        $��SendLaButton��5l   �   
�������#��[ 68   $
��Left=�6_   .��5m      �� $	��220��6%   $��LeftToRight=�61   $ 
Horizontal=�6M   .+�5n        $ SendOutline=�66   ��5o   "   u� 5p   &   $ 20Q�60   6<   5r      �� 5s      $��Send LA��6$   $	��Top=�66   $	��-10q�6N   #���  6\   ��5t   "   u� 5u   &   $ 10Q�60   $ -2Q�6:   6E   5w      5x      5y      �� 5z      
�������#��[ 67   $
��Left=�6^   .��5{      �� $	��220��6%   $��LeftToRight=�61   $ 
Horizontal=�6M   .+�5|        $ SendLock=�66   ��5}   "   u� 5~   &   $ 20Q�60   6<   5�      �� 5�      $��Send Sailing Confirmation��6$   $	��Top=�6H   $	��-10q�6`   #���  6n      Hint$7 1Message Template for sending Sailing Confirmation��6�   ��5�   "   u� 5�   &   $ 10Q�60   $ -2Q�6:   $	
 -10Q�6C   6Q   $	��220��5�   C   5�      5�      5�      5�      5�      5�      5�      ]� 5�   
   -   5        $��	BodyPanel��5�   a   $��Stretch=�6   $��Stretch=�60   
���������5�      u� 5�      $	 250Q�6   6)   5�      5�      5�      