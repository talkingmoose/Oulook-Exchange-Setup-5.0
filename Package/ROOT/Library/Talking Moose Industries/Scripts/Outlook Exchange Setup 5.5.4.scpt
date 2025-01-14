FasdUAS 1.101.10   ��   ��    k             l      ��  ��   ��

--------------------------------------------
Outlook Exchange Setup 5
� Copyright 2008-2018 William Smith
bill@talkingmoose.net

Except where otherwise noted, this work is licensed under
http://creativecommons.org/licenses/by/4.0/

This file is one of four files for assisting a user with configuring
an Exchange account in Microsoft Outlook 2016 for Mac:

1. Outlook Exchange Setup 5.5.3.scpt
2. OutlookExchangeSetupLaunchAgent.sh
3. net.talkingmoose.OutlookExchangeSetupLaunchAgent.plist
4. com.microsoft.Outlook.plist for creating a configuraiton profile

These scripts and files may be freely modified for personal or commercial
purposes but may not be republished for profit without prior consent.

If you find these resources useful or have ideas for improving them,
please let me know. It is only compatible with Outlook 2016 for Mac.

--------------------------------------------

This script assists a user with the setup of his Exchange account
information. Below are basic instructions for using the script.
Consult the Outlook Exchange Setup 5 Administrator's Guide
for complete details.

1.	Customize the "network and  server properties" below with information
	appropriate to your network.
	
2.	Deploy this script to a location on your Macs such as
	"/Library/CompanyName/Outlook Exchange Setup 5.5.2.scpt".

3. 	Deploy the recommended "Outlook preferences.mobileconfig"
	configuration profile to eliminate Outlook's startup windows.
	This assumes you're using the volume license edition
	of Office 2016 for Mac.
	
4.	Deploy the OutlookExchangeSetup5.plist file to
	/Library/LaunchAgents. Update the path to point to the
	OutlookExchangeSetup5.5.2.scpt script.
	  
This script assumes the user's full name is in the form of "Last, First",
but is easily modified if the full name is in the form of "First Last".
It works especially well if the Mac is bound to Active Directory where
the user's short name will match his login name. Optionally, you cans set dscl
to pull the user's mail from a directory service.

     � 	 	� 
 
 - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
 O u t l o o k   E x c h a n g e   S e t u p   5 
 �   C o p y r i g h t   2 0 0 8 - 2 0 1 8   W i l l i a m   S m i t h 
 b i l l @ t a l k i n g m o o s e . n e t 
 
 E x c e p t   w h e r e   o t h e r w i s e   n o t e d ,   t h i s   w o r k   i s   l i c e n s e d   u n d e r 
 h t t p : / / c r e a t i v e c o m m o n s . o r g / l i c e n s e s / b y / 4 . 0 / 
 
 T h i s   f i l e   i s   o n e   o f   f o u r   f i l e s   f o r   a s s i s t i n g   a   u s e r   w i t h   c o n f i g u r i n g 
 a n   E x c h a n g e   a c c o u n t   i n   M i c r o s o f t   O u t l o o k   2 0 1 6   f o r   M a c : 
 
 1 .   O u t l o o k   E x c h a n g e   S e t u p   5 . 5 . 3 . s c p t 
 2 .   O u t l o o k E x c h a n g e S e t u p L a u n c h A g e n t . s h 
 3 .   n e t . t a l k i n g m o o s e . O u t l o o k E x c h a n g e S e t u p L a u n c h A g e n t . p l i s t 
 4 .   c o m . m i c r o s o f t . O u t l o o k . p l i s t   f o r   c r e a t i n g   a   c o n f i g u r a i t o n   p r o f i l e 
 
 T h e s e   s c r i p t s   a n d   f i l e s   m a y   b e   f r e e l y   m o d i f i e d   f o r   p e r s o n a l   o r   c o m m e r c i a l 
 p u r p o s e s   b u t   m a y   n o t   b e   r e p u b l i s h e d   f o r   p r o f i t   w i t h o u t   p r i o r   c o n s e n t . 
 
 I f   y o u   f i n d   t h e s e   r e s o u r c e s   u s e f u l   o r   h a v e   i d e a s   f o r   i m p r o v i n g   t h e m , 
 p l e a s e   l e t   m e   k n o w .   I t   i s   o n l y   c o m p a t i b l e   w i t h   O u t l o o k   2 0 1 6   f o r   M a c . 
 
 - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
 
 T h i s   s c r i p t   a s s i s t s   a   u s e r   w i t h   t h e   s e t u p   o f   h i s   E x c h a n g e   a c c o u n t 
 i n f o r m a t i o n .   B e l o w   a r e   b a s i c   i n s t r u c t i o n s   f o r   u s i n g   t h e   s c r i p t . 
 C o n s u l t   t h e   O u t l o o k   E x c h a n g e   S e t u p   5   A d m i n i s t r a t o r ' s   G u i d e 
 f o r   c o m p l e t e   d e t a i l s . 
 
 1 . 	 C u s t o m i z e   t h e   " n e t w o r k   a n d     s e r v e r   p r o p e r t i e s "   b e l o w   w i t h   i n f o r m a t i o n 
 	 a p p r o p r i a t e   t o   y o u r   n e t w o r k . 
 	 
 2 . 	 D e p l o y   t h i s   s c r i p t   t o   a   l o c a t i o n   o n   y o u r   M a c s   s u c h   a s 
 	 " / L i b r a r y / C o m p a n y N a m e / O u t l o o k   E x c h a n g e   S e t u p   5 . 5 . 2 . s c p t " . 
 
 3 .   	 D e p l o y   t h e   r e c o m m e n d e d   " O u t l o o k   p r e f e r e n c e s . m o b i l e c o n f i g " 
 	 c o n f i g u r a t i o n   p r o f i l e   t o   e l i m i n a t e   O u t l o o k ' s   s t a r t u p   w i n d o w s . 
 	 T h i s   a s s u m e s   y o u ' r e   u s i n g   t h e   v o l u m e   l i c e n s e   e d i t i o n 
 	 o f   O f f i c e   2 0 1 6   f o r   M a c . 
 	 
 4 . 	 D e p l o y   t h e   O u t l o o k E x c h a n g e S e t u p 5 . p l i s t   f i l e   t o 
 	 / L i b r a r y / L a u n c h A g e n t s .   U p d a t e   t h e   p a t h   t o   p o i n t   t o   t h e 
 	 O u t l o o k E x c h a n g e S e t u p 5 . 5 . 2 . s c p t   s c r i p t . 
 	     
 T h i s   s c r i p t   a s s u m e s   t h e   u s e r ' s   f u l l   n a m e   i s   i n   t h e   f o r m   o f   " L a s t ,   F i r s t " , 
 b u t   i s   e a s i l y   m o d i f i e d   i f   t h e   f u l l   n a m e   i s   i n   t h e   f o r m   o f   " F i r s t   L a s t " . 
 I t   w o r k s   e s p e c i a l l y   w e l l   i f   t h e   M a c   i s   b o u n d   t o   A c t i v e   D i r e c t o r y   w h e r e 
 t h e   u s e r ' s   s h o r t   n a m e   w i l l   m a t c h   h i s   l o g i n   n a m e .   O p t i o n a l l y ,   y o u   c a n s   s e t   d s c l 
 t o   p u l l   t h e   u s e r ' s   m a i l   f r o m   a   d i r e c t o r y   s e r v i c e . 
 
   
  
 l     ��������  ��  ��        l     ��  ��      global logMesage     �   "   g l o b a l   l o g M e s a g e      l     ��������  ��  ��        l     ��  ��    0 *------------------------------------------     �   T - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -      l     ��  ��    , & Begin network, server and preferences     �   L   B e g i n   n e t w o r k ,   s e r v e r   a n d   p r e f e r e n c e s      l     ��   ��    0 *------------------------------------------      � ! ! T - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -   " # " l     ��������  ��  ��   #  $ % $ l     ��������  ��  ��   %  & ' & l     �� ( )��   ( C =------------- Exchange Server settings ----------------------    ) � * * z - - - - - - - - - - - - -   E x c h a n g e   S e r v e r   s e t t i n g s   - - - - - - - - - - - - - - - - - - - - - - '  + , + l     ��������  ��  ��   ,  - . - j     �� /�� 0 usekerberos useKerberos / m     ��
�� boovtrue .  0 1 0 l     �� 2 3��   2 B < Set this to true only if Macs in your environment are bound    3 � 4 4 x   S e t   t h i s   t o   t r u e   o n l y   i f   M a c s   i n   y o u r   e n v i r o n m e n t   a r e   b o u n d 1  5 6 5 l     �� 7 8��   7 C = to Active Directory and your network is properly configured.    8 � 9 9 z   t o   A c t i v e   D i r e c t o r y   a n d   y o u r   n e t w o r k   i s   p r o p e r l y   c o n f i g u r e d . 6  : ; : l     ��������  ��  ��   ;  < = < j    �� >��  0 exchangeserver ExchangeServer > m     ? ? � @ @ ( e x c h a n g e . e x a m p l e . c o m =  A B A l     �� C D��   C 6 0 Address of your organization's Exchange server.    D � E E `   A d d r e s s   o f   y o u r   o r g a n i z a t i o n ' s   E x c h a n g e   s e r v e r . B  F G F l     ��������  ��  ��   G  H I H j    �� J�� 60 exchangeserverrequiresssl ExchangeServerRequiresSSL J m    ��
�� boovtrue I  K L K l     �� M N��   M   True for most servers.    N � O O .   T r u e   f o r   m o s t   s e r v e r s . L  P Q P l     ��������  ��  ��   Q  R S R j   	 �� T�� .0 exchangeserversslport ExchangeServerSSLPort T m   	 
����� S  U V U l     �� W X��   W @ : If ExchangeServerRequiresSSL is true set the port to 443.    X � Y Y t   I f   E x c h a n g e S e r v e r R e q u i r e s S S L   i s   t r u e   s e t   t h e   p o r t   t o   4 4 3 . V  Z [ Z l     �� \ ]��   \ @ : If ExchangeServerRequiresSSL is false set the port to 80.    ] � ^ ^ t   I f   E x c h a n g e S e r v e r R e q u i r e s S S L   i s   f a l s e   s e t   t h e   p o r t   t o   8 0 . [  _ ` _ l     �� a b��   a L F Use a different port number only if your administrator instructs you.    b � c c �   U s e   a   d i f f e r e n t   p o r t   n u m b e r   o n l y   i f   y o u r   a d m i n i s t r a t o r   i n s t r u c t s   y o u . `  d e d l     ��������  ��  ��   e  f g f j    �� h�� "0 directoryserver DirectoryServer h m     i i � j j  g c . e x a m p l e . c o m g  k l k l     �� m n��   m Z T Address of an internal Global Catalog server (a type of Windows domain controller).    n � o o �   A d d r e s s   o f   a n   i n t e r n a l   G l o b a l   C a t a l o g   s e r v e r   ( a   t y p e   o f   W i n d o w s   d o m a i n   c o n t r o l l e r ) . l  p q p l     �� r s��   r L F The LDAP server in a Windows network will be a Global Catalog server,    s � t t �   T h e   L D A P   s e r v e r   i n   a   W i n d o w s   n e t w o r k   w i l l   b e   a   G l o b a l   C a t a l o g   s e r v e r , q  u v u l     �� w x��   w 2 , which is separate from the Exchange Server.    x � y y X   w h i c h   i s   s e p a r a t e   f r o m   t h e   E x c h a n g e   S e r v e r . v  z { z l     ��������  ��  ��   {  | } | j    �� ~�� N0 %directoryserverrequiresauthentication %DirectoryServerRequiresAuthentication ~ m    ��
�� boovtrue }   �  l     �� � ���   � ' ! This will almost always be true.    � � � � B   T h i s   w i l l   a l m o s t   a l w a y s   b e   t r u e . �  � � � l     ��������  ��  ��   �  � � � j    �� ��� 80 directoryserverrequiresssl DirectoryServerRequiresSSL � m    ��
�� boovtrue �  � � � l     �� � ���   � ' ! This will almost always be true.    � � � � B   T h i s   w i l l   a l m o s t   a l w a y s   b e   t r u e . �  � � � l     ��������  ��  ��   �  � � � j    �� ��� 00 directoryserversslport DirectoryServerSSLPort � m    ����� �  � � � l     �� � ���   � B < If DirectoryServerRequiresSSL is true set the port to 3269.    � � � � x   I f   D i r e c t o r y S e r v e r R e q u i r e s S S L   i s   t r u e   s e t   t h e   p o r t   t o   3 2 6 9 . �  � � � l     �� � ���   � C = If DirectoryServerRequiresSSL is false set the port to 3268.    � � � � z   I f   D i r e c t o r y S e r v e r R e q u i r e s S S L   i s   f a l s e   s e t   t h e   p o r t   t o   3 2 6 8 . �  � � � l     �� � ���   � U O Use a different port number only if your Exchange administrator instructs you.    � � � � �   U s e   a   d i f f e r e n t   p o r t   n u m b e r   o n l y   i f   y o u r   E x c h a n g e   a d m i n i s t r a t o r   i n s t r u c t s   y o u . �  � � � l     ��������  ��  ��   �  � � � j    �� ��� >0 directoryservermaximumresults DirectoryServerMaximumResults � m    ����� �  � � � l     �� � ���   � G A When searching the Global Catalog server, this number determines    � � � � �   W h e n   s e a r c h i n g   t h e   G l o b a l   C a t a l o g   s e r v e r ,   t h i s   n u m b e r   d e t e r m i n e s �  � � � l     �� � ���   � 0 * the maximum number of entries to display.    � � � � T   t h e   m a x i m u m   n u m b e r   o f   e n t r i e s   t o   d i s p l a y . �  � � � l     ��������  ��  ��   �  � � � j    �� ��� 60 directoryserversearchbase DirectoryServerSearchBase � m     � � � � �   �  � � � l     �� � ���   � + % example: "cn=users,dc=domain,dc=com"    � � � � J   e x a m p l e :   " c n = u s e r s , d c = d o m a i n , d c = c o m " �  � � � l     �� � ���   �   Usually, this is empty.    � � � � 0   U s u a l l y ,   t h i s   i s   e m p t y . �  � � � l     ��������  ��  ��   �  � � � l     ��������  ��  ��   �  � � � l     �� � ���   � D >------------- For Active Directory users ---------------------    � � � � | - - - - - - - - - - - - -   F o r   A c t i v e   D i r e c t o r y   u s e r s   - - - - - - - - - - - - - - - - - - - - - �  � � � l     ��������  ��  ��   �  � � � j     �� ��� N0 %getuserinformationfromactivedirectory %getUserInformationFromActiveDirectory � m    ��
�� boovtrue �  � � � l     �� � ���   � N H If Macs are bound to Active Directory via dsconfigad/Directory Utility,    � � � � �   I f   M a c s   a r e   b o u n d   t o   A c t i v e   D i r e c t o r y   v i a   d s c o n f i g a d / D i r e c t o r y   U t i l i t y , �  � � � l     �� � ���   � ^ X they can use dscl to return the current user's email address, phone number, title, etc.    � � � � �   t h e y   c a n   u s e   d s c l   t o   r e t u r n   t h e   c u r r e n t   u s e r ' s   e m a i l   a d d r e s s ,   p h o n e   n u m b e r ,   t i t l e ,   e t c . �  � � � l     ��������  ��  ��   �  � � � l     ��������  ��  ��   �  � � � l     �� � ���   � > 8------------- For Office 365 users ---------------------    � � � � p - - - - - - - - - - - - -   F o r   O f f i c e   3 6 5   u s e r s   - - - - - - - - - - - - - - - - - - - - - �  � � � l     ��������  ��  ��   �  � � � j   ! #�� ��� *0 useemailforusername useEmailForUsername � m   ! "��
�� boovfals �  � � � l     �� � ���   � S M Office 365 and similar mail services may require the username to be the same    � � � � �   O f f i c e   3 6 5   a n d   s i m i l a r   m a i l   s e r v i c e s   m a y   r e q u i r e   t h e   u s e r n a m e   t o   b e   t h e   s a m e �  � � � l     �� � ���   � P J as the email address. Set this to true if the username is the same as the    � � � � �   a s   t h e   e m a i l   a d d r e s s .   S e t   t h i s   t o   t r u e   i f   t h e   u s e r n a m e   i s   t h e   s a m e   a s   t h e �  � � � l     �� � ���   �   email address.    � � � �    e m a i l   a d d r e s s . �  � � � l     ��������  ��  ��   �  � � � l     ��������  ��  ��   �  � � � l     �� � ���   � B <------------- For non Active Directory users ---------------    � � � � x - - - - - - - - - - - - -   F o r   n o n   A c t i v e   D i r e c t o r y   u s e r s   - - - - - - - - - - - - - - - �    l     ��������  ��  ��    j   $ (���� 0 
domainname 
domainName m   $ ' �  e x a m p l e . c o m  l     ��	
��  	 P J Complete this only if not using Active Directory to get user information.   
 � �   C o m p l e t e   t h i s   o n l y   i f   n o t   u s i n g   A c t i v e   D i r e c t o r y   t o   g e t   u s e r   i n f o r m a t i o n .  l     ����   L F The part of your organization's email address following the @ symbol.    � �   T h e   p a r t   o f   y o u r   o r g a n i z a t i o n ' s   e m a i l   a d d r e s s   f o l l o w i n g   t h e   @   s y m b o l .  l     ����~��  �  �~    j   ) +�}�} 0 emailformat emailFormat m   ) *�|�|   l     �{�{   P J Complete this only if not using Active Directory to get user information.    � �   C o m p l e t e   t h i s   o n l y   i f   n o t   u s i n g   A c t i v e   D i r e c t o r y   t o   g e t   u s e r   i n f o r m a t i o n .  l     �z�z   P J When Active Directory is unavailable to determine a user's email address,    � �   W h e n   A c t i v e   D i r e c t o r y   i s   u n a v a i l a b l e   t o   d e t e r m i n e   a   u s e r ' s   e m a i l   a d d r e s s ,  !  l     �y"#�y  " V P this script will attempt to parse it from the display name of the user's login.   # �$$ �   t h i s   s c r i p t   w i l l   a t t e m p t   t o   p a r s e   i t   f r o m   t h e   d i s p l a y   n a m e   o f   t h e   u s e r ' s   l o g i n .! %&% l     �x�w�v�x  �w  �v  & '(' l     �u)*�u  ) 1 + Describe your organization's email format:   * �++ V   D e s c r i b e   y o u r   o r g a n i z a t i o n ' s   e m a i l   f o r m a t :( ,-, l     �t./�t  . / ) 1: Email format is first.last@domain.com   / �00 R   1 :   E m a i l   f o r m a t   i s   f i r s t . l a s t @ d o m a i n . c o m- 121 l     �s34�s  3 * $ 2: Email format is first@domain.com   4 �55 H   2 :   E m a i l   f o r m a t   i s   f i r s t @ d o m a i n . c o m2 676 l     �r89�r  8 N H 3: Email format is flast@domain.com (first name initial plus last name)   9 �:: �   3 :   E m a i l   f o r m a t   i s   f l a s t @ d o m a i n . c o m   ( f i r s t   n a m e   i n i t i a l   p l u s   l a s t   n a m e )7 ;<; l     �q=>�q  = . ( 4: Email format is shortName@domain.com   > �?? P   4 :   E m a i l   f o r m a t   i s   s h o r t N a m e @ d o m a i n . c o m< @A@ l     �p�o�n�p  �o  �n  A BCB j   , .�mD�m 0 displayname displayNameD m   , -�l�l C EFE l     �kGH�k  G P J Complete this only if not using Active Directory to get user information.   H �II �   C o m p l e t e   t h i s   o n l y   i f   n o t   u s i n g   A c t i v e   D i r e c t o r y   t o   g e t   u s e r   i n f o r m a t i o n .F JKJ l     �jLM�j  L M G Describe how the user's display name appears at the bottom of the menu   M �NN �   D e s c r i b e   h o w   t h e   u s e r ' s   d i s p l a y   n a m e   a p p e a r s   a t   t h e   b o t t o m   o f   t h e   m e n uK OPO l     �iQR�i  Q R L when clicking the Apple menu (Log Out Joe Cool... or Log Out Cool, Joe...).   R �SS �   w h e n   c l i c k i n g   t h e   A p p l e   m e n u   ( L o g   O u t   J o e   C o o l . . .   o r   L o g   O u t   C o o l ,   J o e . . . ) .P TUT l     �hVW�h  V / ) 1: Display name appears as "Last, First"   W �XX R   1 :   D i s p l a y   n a m e   a p p e a r s   a s   " L a s t ,   F i r s t "U YZY l     �g[\�g  [ . ( 2: Display name appears as "First Last"   \ �]] P   2 :   D i s p l a y   n a m e   a p p e a r s   a s   " F i r s t   L a s t "Z ^_^ l     �f�e�d�f  �e  �d  _ `a` j   / 3�cb�c 0 domainprefix domainPrefixb m   / 2cc �dd  a efe l     �bgh�b  g Y S Optionally append a NetBIOS domain name to the beginning of the user's short name.   h �ii �   O p t i o n a l l y   a p p e n d   a   N e t B I O S   d o m a i n   n a m e   t o   t h e   b e g i n n i n g   o f   t h e   u s e r ' s   s h o r t   n a m e .f jkj l     �alm�a  l 9 3 Be sure to use two backslashes when adding a name.   m �nn f   B e   s u r e   t o   u s e   t w o   b a c k s l a s h e s   w h e n   a d d i n g   a   n a m e .k opo l     �`qr�`  q N H Example: Use "TALKINGMOOSE\\" to set user name "TALKINGMOOSE\username".   r �ss �   E x a m p l e :   U s e   " T A L K I N G M O O S E \ \ "   t o   s e t   u s e r   n a m e   " T A L K I N G M O O S E \ u s e r n a m e " .p tut l     �_�^�]�_  �^  �]  u vwv l     �\�[�Z�\  �[  �Z  w xyx l     �Yz{�Y  z C =------------- User Experience -------------------------------   { �|| z - - - - - - - - - - - - -   U s e r   E x p e r i e n c e   - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -y }~} l     �X�W�V�X  �W  �V  ~ � j   4 8�U��U (0 verifyemailaddress verifyEMailAddress� m   4 5�T
�T boovfals� ��� l     �S���S  � M G If set to "true", a dialog asks the user to confirm his email address.   � ��� �   I f   s e t   t o   " t r u e " ,   a   d i a l o g   a s k s   t h e   u s e r   t o   c o n f i r m   h i s   e m a i l   a d d r e s s .� ��� l     �R�Q�P�R  �Q  �P  � ��� j   9 =�O��O *0 verifyserveraddress verifyServerAddress� m   9 :�N
�N boovfals� ��� l     �M���M  � W Q If set to "true", a dialog asks the user to confirm his Exchange server address.   � ��� �   I f   s e t   t o   " t r u e " ,   a   d i a l o g   a s k s   t h e   u s e r   t o   c o n f i r m   h i s   E x c h a n g e   s e r v e r   a d d r e s s .� ��� l     �L�K�J�L  �K  �J  � ��� j   > B�I��I *0 displaydomainprefix displayDomainPrefix� m   > ?�H
�H boovfals� ��� l     �G���G  � C = If set to "true", the username appears as "DOMAIN\username".   � ��� z   I f   s e t   t o   " t r u e " ,   t h e   u s e r n a m e   a p p e a r s   a s   " D O M A I N \ u s e r n a m e " .� ��� l     �F���F  � 5 / Otherwise, the username appears as "username".   � ��� ^   O t h e r w i s e ,   t h e   u s e r n a m e   a p p e a r s   a s   " u s e r n a m e " .� ��� l     �E�D�C�E  �D  �C  � ��� j   C G�B��B *0 downloadheadersonly downloadHeadersOnly� m   C D�A
�A boovfals� ��� l     �@���@  � H B If set to "true", only email headers are downloaded into Outlook.   � ��� �   I f   s e t   t o   " t r u e " ,   o n l y   e m a i l   h e a d e r s   a r e   d o w n l o a d e d   i n t o   O u t l o o k .� ��� l     �?���?  � B < This takes much less time to sync but a user must be online   � ��� x   T h i s   t a k e s   m u c h   l e s s   t i m e   t o   s y n c   b u t   a   u s e r   m u s t   b e   o n l i n e� ��� l     �>���>  � %  to download and view messages.   � ��� >   t o   d o w n l o a d   a n d   v i e w   m e s s a g e s .� ��� l     �=�<�;�=  �<  �;  � ��� j   H L�:��: 20 hideonmycomputerfolders hideOnMyComputerFolders� m   H I�9
�9 boovfals� ��� l     �8���8  � - ' If set to "true", hides local folders.   � ��� N   I f   s e t   t o   " t r u e " ,   h i d e s   l o c a l   f o l d e r s .� ��� l     �7���7  � ; 5 A single Exchange account should do this by default.   � ��� j   A   s i n g l e   E x c h a n g e   a c c o u n t   s h o u l d   d o   t h i s   b y   d e f a u l t .� ��� l     �6�5�4�6  �5  �4  � ��� j   M Q�3��3 0 unifiedinbox unifiedInbox� m   M N�2
�2 boovfals� ��� l     �1���1  � C = If set to "true", turns on the Group Similar Folders feature   � ��� z   I f   s e t   t o   " t r u e " ,   t u r n s   o n   t h e   G r o u p   S i m i l a r   F o l d e r s   f e a t u r e� ��� l     �0���0  � / ) in Outlook menu > Preferences > General.   � ��� R   i n   O u t l o o k   m e n u   >   P r e f e r e n c e s   >   G e n e r a l .� ��� l     �/�.�-�/  �.  �-  � ��� j   R V�,��, (0 enableautodiscover enableAutodiscover� m   R S�+
�+ boovtrue� ��� l     �*���*  � < 6 If set to "true", disables Autodiscover functionality   � ��� l   I f   s e t   t o   " t r u e " ,   d i s a b l e s   A u t o d i s c o v e r   f u n c t i o n a l i t y� ��� l     �)���)  � C = for the Exchange account. Not recommended for mobile devices   � ��� z   f o r   t h e   E x c h a n g e   a c c o u n t .   N o t   r e c o m m e n d e d   f o r   m o b i l e   d e v i c e s� ��� l     �(���(  � B < that may connect to an internal Exchange server address and   � ��� x   t h a t   m a y   c o n n e c t   t o   a n   i n t e r n a l   E x c h a n g e   s e r v e r   a d d r e s s   a n d� ��� l     �'���'  � ? 9 connect to a different external Exchange server address.   � ��� r   c o n n e c t   t o   a   d i f f e r e n t   e x t e r n a l   E x c h a n g e   s e r v e r   a d d r e s s .� ��� l     �&�%�$�&  �%  �$  � ��� j   W ]�#��# 0 errormessage errorMessage� m   W Z�� ��� � O u t l o o k ' s   s e t u p   f o r   y o u r   E x c h a n g e   a c c o u n t   f a i l e d .   P l e a s e   c o n t a c t   t h e   H e l p   D e s k   f o r   a s s i s t a n c e .� ��� l     �"���"  � T N Customize this error message for your users in case their account setup fails   � ��� �   C u s t o m i z e   t h i s   e r r o r   m e s s a g e   f o r   y o u r   u s e r s   i n   c a s e   t h e i r   a c c o u n t   s e t u p   f a i l s� ��� l     �!� ��!  �   �  � ��� l     ����  � 0 *------------------------------------------   � ��� T - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -� ��� l     � �    * $ End network, server and preferences    � H   E n d   n e t w o r k ,   s e r v e r   a n d   p r e f e r e n c e s�  l     ��   0 *------------------------------------------    � T - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 	 l     ����  �  �  	 

 l     ��   0 *------------------------------------------    � T - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -  l     ��     Begin log file setup    � *   B e g i n   l o g   f i l e   s e t u p  l     ��   0 *------------------------------------------    � T - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -  l     ����  �  �    l     ��   < 6 create the log file in the current user's Logs folder    � l   c r e a t e   t h e   l o g   f i l e   i n   t h e   c u r r e n t   u s e r ' s   L o g s   f o l d e r  !  l     ����  �  �  ! "#" l    $��$ I     �%�� 0 writelog writeLog% &�
& m    '' �(( D S t a r t i n g   E x c h a n g e   a c c o u n t   s e t u p . . .�
  �  �  �  # )*) l   +�	�+ I    �,�� 0 writelog writeLog, -�- b    ./. m    	00 �11  S c r i p t :  / n   	 232 1   
 �
� 
pnam3  f   	 
�  �  �	  �  * 454 l   6��6 I    �7� � 0 writelog writeLog7 8��8 o    ��
�� 
ret ��  �   �  �  5 9:9 l     ��������  ��  ��  : ;<; l     ��=>��  = 0 *------------------------------------------   > �?? T - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -< @A@ l     ��BC��  B   End log file setup    C �DD (   E n d   l o g   f i l e   s e t u p  A EFE l     ��GH��  G 0 *------------------------------------------   H �II T - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -F JKJ l     ��������  ��  ��  K LML l     ��NO��  N 0 *------------------------------------------   O �PP T - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -M QRQ l     ��ST��  S &   Begin logging script properties   T �UU @   B e g i n   l o g g i n g   s c r i p t   p r o p e r t i e sR VWV l     ��XY��  X 0 *------------------------------------------   Y �ZZ T - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -W [\[ l     ��������  ��  ��  \ ]^] l   _����_ I    ��`���� 0 writelog writeLog` a��a m    bb �cc & S e t u p   p r o p e r t i e s . . .��  ��  ��  ��  ^ ded l    ,f����f I     ,��g���� 0 writelog writeLogg h��h b   ! (iji m   ! "kk �ll  U s e   K e r b e r o s :  j o   " '���� 0 usekerberos useKerberos��  ��  ��  ��  e mnm l  - 9o����o I   - 9��p���� 0 writelog writeLogp q��q b   . 5rsr m   . /tt �uu " E x c h a n g e   S e r v e r :  s o   / 4����  0 exchangeserver ExchangeServer��  ��  ��  ��  n vwv l  : Fx����x I   : F��y���� 0 writelog writeLogy z��z b   ; B{|{ m   ; <}} �~~ < E x c h a n g e   S e r v e r   R e q u i r e s   S S L :  | o   < A���� 60 exchangeserverrequiresssl ExchangeServerRequiresSSL��  ��  ��  ��  w � l  G S������ I   G S������� 0 writelog writeLog� ���� b   H O��� m   H I�� ��� , E x c h a n g e   S e r v e r   P o r t :  � o   I N���� .0 exchangeserversslport ExchangeServerSSLPort��  ��  ��  ��  � ��� l  T `������ I   T `������� 0 writelog writeLog� ���� b   U \��� m   U V�� ��� $ D i r e c t o r y   S e r v e r :  � o   V [���� "0 directoryserver DirectoryServer��  ��  ��  ��  � ��� l  a m������ I   a m������� 0 writelog writeLog� ���� b   b i��� m   b c�� ��� T D i r e c t o r y   S e r v e r   R e q u i r e s   A u t h e n t i c a t i o n :  � o   c h���� N0 %directoryserverrequiresauthentication %DirectoryServerRequiresAuthentication��  ��  ��  ��  � ��� l  n z������ I   n z������� 0 writelog writeLog� ���� b   o v��� m   o p�� ��� > D i r e c t o r y   S e r v e r   R e q u i r e s   S S L :  � o   p u���� 80 directoryserverrequiresssl DirectoryServerRequiresSSL��  ��  ��  ��  � ��� l  { ������� I   { �������� 0 writelog writeLog� ���� b   | ���� m   | }�� ��� 6 D i r e c t o r y   S e r v e r   S S L   P o r t :  � o   } ����� 00 directoryserversslport DirectoryServerSSLPort��  ��  ��  ��  � ��� l  � ������� I   � �������� 0 writelog writeLog� ���� b   � ���� m   � ��� ��� D D i r e c t o r y   S e r v e r   M a x i m u m   R e s u l t s :  � o   � ����� >0 directoryservermaximumresults DirectoryServerMaximumResults��  ��  ��  ��  � ��� l  � ������� I   � �������� 0 writelog writeLog� ���� b   � ���� m   � ��� ��� < D i r e c t o r y   S e r v e r   S e a r c h   B a s e :  � o   � ����� 60 directoryserversearchbase DirectoryServerSearchBase��  ��  ��  ��  � ��� l  � ������� I   � �������� 0 writelog writeLog� ���� b   � ���� m   � ��� ��� X G e t   U s e r   I n f o r m a t i o n   f r o m   A c t i v e   D i r e c t o r y :  � o   � ����� N0 %getuserinformationfromactivedirectory %getUserInformationFromActiveDirectory��  ��  ��  ��  � ��� l  � ������� I   � �������� 0 writelog writeLog� ���� o   � ���
�� 
ret ��  ��  ��  ��  � ��� l     ��������  ��  ��  � ��� l  ������� Z   �������� =  � ���� o   � ����� N0 %getuserinformationfromactivedirectory %getUserInformationFromActiveDirectory� m   � ���
�� boovfals� k   ��� ��� I   � �������� 0 writelog writeLog� ���� b   � ���� m   � ��� ���  D o m a i n   N a m e :  � o   � ����� 0 
domainname 
domainName��  ��  � ��� I   � �������� 0 writelog writeLog� ���� b   � ���� m   � ��� ���  E m a i l   f o r m a t :  � o   � ����� 0 emailformat emailFormat��  ��  � ��� I   � �������� 0 writelog writeLog� ���� b   � ���� m   � ��� ���  D i s p l a y   N a m e :  � o   � ����� 0 displayname displayName��  ��  � ��� I   � �������� 0 writelog writeLog� ���� b   � ���� m   � ��� ���  D o m a i n   P r e f i x :  � o   � ����� 0 domainprefix domainPrefix��  ��  � ���� I   �������� 0 writelog writeLog� ���� o   � ��
�� 
ret ��  ��  ��  ��  ��  ��  ��  � ��� l     ��������  ��  ��  � ��� l 	����� I  	�~��}�~ 0 writelog writeLog� ��|� b  
   m  
 � , V e r i f y   E m a i l   A d d r e s s :   o  �{�{ (0 verifyemailaddress verifyEMailAddress�|  �}  ��  �  �  l &�z�y I  &�x�w�x 0 writelog writeLog �v b  "	
	 m   � . V e r i f y   S e r v e r   A d d r e s s :  
 o  !�u�u *0 verifyserveraddress verifyServerAddress�v  �w  �z  �y    l '5�t�s I  '5�r�q�r 0 writelog writeLog �p b  (1 m  (+ � . D i s p l a y   D o m a i n   P r e f i x :   o  +0�o�o *0 displaydomainprefix displayDomainPrefix�p  �q  �t  �s    l 6D�n�m I  6D�l�k�l 0 writelog writeLog �j b  7@ m  7: � . D o w n l o a d   H e a d e r s   O n l y :   o  :?�i�i *0 downloadheadersonly downloadHeadersOnly�j  �k  �n  �m     l ES!�h�g! I  ES�f"�e�f 0 writelog writeLog" #�d# b  FO$%$ m  FI&& �'' : H i d e   O n   M y   C o m p u t e r   F o l d e r s :  % o  IN�c�c 20 hideonmycomputerfolders hideOnMyComputerFolders�d  �e  �h  �g    ()( l Tb*�b�a* I  Tb�`+�_�` 0 writelog writeLog+ ,�^, b  U^-.- m  UX// �00  U n i f i e d   I n b o x :  . o  X]�]�] 0 unifiedinbox unifiedInbox�^  �_  �b  �a  ) 121 l cq3�\�[3 I  cq�Z4�Y�Z 0 writelog writeLog4 5�X5 b  dm676 m  dg88 �99 , D i s a b l e   A u t o d i s c o v e r :  7 o  gl�W�W (0 enableautodiscover enableAutodiscover�X  �Y  �\  �[  2 :;: l r�<�V�U< I  r��T=�S�T 0 writelog writeLog= >�R> b  s|?@? m  svAA �BB ( E r r o r   M e s s a g e   t e x t :  @ o  v{�Q�Q 0 errormessage errorMessage�R  �S  �V  �U  ; CDC l ��E�P�OE I  ���NF�M�N 0 writelog writeLogF G�LG o  ���K
�K 
ret �L  �M  �P  �O  D HIH l     �J�I�H�J  �I  �H  I JKJ l     �GLM�G  L 0 *------------------------------------------   M �NN T - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -K OPO l     �FQR�F  Q %  End logging script properties    R �SS >   E n d   l o g g i n g   s c r i p t   p r o p e r t i e s  P TUT l     �EVW�E  V 0 *------------------------------------------   W �XX T - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -U YZY l     �D�C�B�D  �C  �B  Z [\[ l     �A]^�A  ] 0 *------------------------------------------   ^ �__ T - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -\ `a` l     �@bc�@  b ( " Begin collecting user information   c �dd D   B e g i n   c o l l e c t i n g   u s e r   i n f o r m a t i o na efe l     �?gh�?  g 0 *------------------------------------------   h �ii T - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -f jkj l     �>�=�<�>  �=  �<  k lml l     �;no�;  n R L attempt to read information from Active Directory for the Me Contact record   o �pp �   a t t e m p t   t o   r e a d   i n f o r m a t i o n   f r o m   A c t i v e   D i r e c t o r y   f o r   t h e   M e   C o n t a c t   r e c o r dm qrq l     �:�9�8�:  �9  �8  r sts l ��u�7�6u r  ��vwv m  ��xx �yy  w o      �5�5 0 userfirstname userFirstName�7  �6  t z{z l ��|�4�3| r  ��}~} m  �� ���  ~ o      �2�2 0 userlastname userLastName�4  �3  { ��� l ����1�0� r  ����� m  ���� ���  � o      �/�/  0 userdepartment userDepartment�1  �0  � ��� l ����.�-� r  ����� m  ���� ���  � o      �,�, 0 
useroffice 
userOffice�.  �-  � ��� l ����+�*� r  ����� m  ���� ���  � o      �)�) 0 usercompany userCompany�+  �*  � ��� l ����(�'� r  ����� m  ���� ���  � o      �&�& 0 userworkphone userWorkPhone�(  �'  � ��� l ����%�$� r  ����� m  ���� ���  � o      �#�# 0 
usermobile 
userMobile�%  �$  � ��� l ����"�!� r  ����� m  ���� ���  � o      � �  0 userfax userFax�"  �!  � ��� l ������ r  ����� m  ���� ���  � o      �� 0 	usertitle 	userTitle�  �  � ��� l ������ r  ����� m  ���� ���  � o      �� 0 
userstreet 
userStreet�  �  � ��� l ������ r  ����� m  ���� ���  � o      �� 0 usercity userCity�  �  � ��� l ������ r  ����� m  ���� ���  � o      �� 0 	userstate 	userState�  �  � ��� l ������ r  ����� m  ���� ���  � o      ��  0 userpostalcode userPostalCode�  �  � ��� l ������ r  ����� m  ���� ���  � o      �� 0 usercountry userCountry�  �  � ��� l ������ r  ����� m  ���� ���  � o      �� 0 userwebpage userWebPage�  �  � ��� l     �
�	��
  �	  �  � ��� l  ����� Z   ������ =  ��� o   �� N0 %getuserinformationfromactivedirectory %getUserInformationFromActiveDirectory� m  �
� boovtrue� k  
:�� ��� l 

����  �  �  � ��� l 

� ���   � + % Get information from Active Directoy   � ��� J   G e t   i n f o r m a t i o n   f r o m   A c t i v e   D i r e c t o y� ��� l 

��������  ��  ��  � ��� l 

������  � 3 - get the domain's primary NetBIOS domain name   � ��� Z   g e t   t h e   d o m a i n ' s   p r i m a r y   N e t B I O S   d o m a i n   n a m e� ��� l 

��������  ��  ��  � ��� Q  
����� k  g�� ��� r     I ����
�� .sysoexecTEXT���     TEXT m   � � / u s r / b i n / d s c l   " / A c t i v e   D i r e c t o r y / "   - r e a d   /   S u b N o d e s   |   a w k   ' B E G I N   { F S = " :   " }   { p r i n t   $ 2 } '��   o      ���� 0 netbiosdomain netbiosDomain�  I  %������ 0 writelog writeLog �� b  !	
	 m   � 0 G e t t i n g   N e t B I O S   d o m a i n :  
 o   ���� 0 netbiosdomain netbiosDomain��  ��   �� Z  &g�� = &- o  &+���� *0 displaydomainprefix displayDomainPrefix m  +,��
�� boovtrue k  0L  r  0= b  07 o  03���� 0 netbiosdomain netbiosDomain m  36 �  \ o      ���� 0 domainprefix domainPrefix �� I  >L������ 0 writelog writeLog �� b  ?H  m  ?B!! �"" . G e t t i n g   d o m a i n   p r e f i x :    o  BG���� 0 domainprefix domainPrefix��  ��  ��  ��   k  Og## $%$ r  OX&'& m  OR(( �))  ' o      ���� 0 domainprefix domainPrefix% *��* I  Yg��+���� 0 writelog writeLog+ ,��, b  Zc-.- m  Z]// �00 . G e t t i n g   d o m a i n   p r e f i x :  . o  ]b���� 0 domainprefix domainPrefix��  ��  ��  ��  � R      ������
�� .ascrerr ****      � ****��  ��  � k  o�11 232 l oo��������  ��  ��  3 454 l oo��67��  6   something went wrong   7 �88 *   s o m e t h i n g   w e n t   w r o n g5 9:9 l oo��������  ��  ��  : ;<; I o���=>
�� .sysodlogaskr        TEXT= b  o|?@? b  oxABA b  ovCDC o  ot���� 0 errormessage errorMessageD o  tu��
�� 
ret B o  vw��
�� 
ret @ m  x{EE �FF � U n a b l e   t o   d e t e r m i n e   N E T B I O S   d o m a i n   n a m e .   T h i s   c o m p u t e r   m a y   n o t   b e   b o u n d   t o   A c t i v e   D i r e c t o r y .> ��GH
�� 
dispG m  ���
�� stic    H ��IJ
�� 
btnsI J  ��KK L��L m  ��MM �NN  O K��  J ��OP
�� 
dfltO J  ��QQ R��R m  ��SS �TT  O K��  P ��U��
�� 
apprU m  ��VV �WW , O u t l o o k   E x c h a n g e   S e t u p��  < XYX I  ����Z���� 0 writelog writeLogZ [��[ m  ��\\ �]] ^ E R R O R :   U n a b l e   t o   d e t e r m i n e   N E T B I O S   d o m a i n   n a m e .��  ��  Y ^��^ R  ������_
�� .ascrerr ****      � ****��  _ ��`��
�� 
errn` m  ����������  ��  � aba l ����������  ��  ��  b cdc l ����ef��  e 7 1 Read full user information from Active Directory   f �gg b   R e a d   f u l l   u s e r   i n f o r m a t i o n   f r o m   A c t i v e   D i r e c t o r yd hih l ����������  ��  ��  i jkj Q  �1lmnl k  ��oo pqp r  ��rsr J  ��tt u��u m  ��vv �ww  :  ��  s n     xyx 1  ����
�� 
txdly 1  ����
�� 
ascrq z{z r  ��|}| I ����~��
�� .sysoexecTEXT���     TEXT~ b  ��� b  ����� m  ���� ��� B / u s r / b i n / d s c l   " / A c t i v e   D i r e c t o r y /� o  ������ 0 netbiosdomain netbiosDomain� m  ���� ���� / A l l   D o m a i n s / "   - r e a d   / U s e r s / $ U S E R   A u t h e n t i c a t i o n A u t h o r i t y   C i t y   c o   c o m p a n y   d e p a r t m e n t   p h y s i c a l D e l i v e r y O f f i c e N a m e   s A M A c c o u n t N a m e   w W W H o m e P a g e   E M a i l A d d r e s s   F A X N u m b e r   F i r s t N a m e   J o b T i t l e   L a s t N a m e   M o b i l e N u m b e r   P h o n e N u m b e r   P o s t a l C o d e   R e a l N a m e   S t a t e   S t r e e t��  } o      ���� "0 userinformation userInformation{ ��� I  ��������� 0 writelog writeLog� ���� b  ����� m  ���� ��� ` G e t t i n g   u s e r   i n f o r m a t i o n   f r o m   A c t i v e   D i r e c t o r y :  � o  ������ "0 userinformation userInformation��  ��  � ���� l ����������  ��  ��  ��  m R      ������
�� .ascrerr ****      � ****��  ��  n k  �1�� ��� l ����������  ��  ��  � ��� l ��������  �   something went wrong   � ��� *   s o m e t h i n g   w e n t   w r o n g� ��� l ����������  ��  ��  � ��� I �����
�� .sysodlogaskr        TEXT� b  ����� b  ����� b  ����� o  ������ 0 errormessage errorMessage� o  ����
�� 
ret � o  ����
�� 
ret � m  ���� ��� n U n a b l e   t o   r e a d   u s e r   i n f o r m a t i o n   f r o m   n e t w o r k   d i r e c t o r y .� ����
�� 
disp� m  ���
�� stic    � ����
�� 
btns� J  	�� ���� m  �� ���  O K��  � ����
�� 
dflt� J  �� ���� m  �� ���  O K��  � �����
�� 
appr� m  �� ��� , O u t l o o k   E x c h a n g e   S e t u p��  � ��� I  &������� 0 writelog writeLog� ���� m  "�� ��� | E R R O R :   U n a b l e   t o   r e a d   u s e r   i n f o r m a t i o n   f r o m   n e t w o r k   d i r e c t o r y .��  ��  � ���� R  '1�����
�� .ascrerr ****      � ****��  � �����
�� 
errn� m  +.��������  ��  k ��� l 22��������  ��  ��  � ��� Y  2
��������� k  F
��� ��� l FF��������  ��  ��  � ��� r  FS��� J  FK�� ���� m  FI�� ���  :  ��  � n     ��� 1  NR��
�� 
txdl� 1  KN��
�� 
ascr� ��� Z  T�������� C T`��� n  T\��� 4  W\��
� 
cpar� o  Z[�~�~ 0 i  � o  TW�}�} "0 userinformation userInformation� m  \_�� ���  E M a i l A d d r e s s :� Q  c����� r  fw��� n  fs��� 4  ns�|�
�| 
citm� m  qr�{�{ � n  fn��� 4  in�z�
�z 
cpar� o  lm�y�y 0 i  � o  fi�x�x "0 userinformation userInformation� o      �w�w 0 emailaddress emailAddress� R      �v�u�t
�v .ascrerr ****      � ****�u  �t  � k  ��� ��� r  ���� J  ��� ��s� m  ��� ���  �s  � n     ��� 1  ���r
�r 
txdl� 1  ���q
�q 
ascr� ��p� r  ����� c  ����� n  ����� 7���o��
�o 
cha � m  ���n�n �  ;  ��� n  ����� 4  ���m 
�m 
cpar  l ���l�k [  �� o  ���j�j 0 i   m  ���i�i �l  �k  � o  ���h�h "0 userinformation userInformation� m  ���g
�g 
TEXT� o      �f�f 0 emailaddress emailAddress�p  ��  ��  �  l ���e�d�c�e  �d  �c    r  ��	 J  ��

 �b m  �� �  :  �b  	 n      1  ���a
�a 
txdl 1  ���`
�` 
ascr  Z  ��_�^ C �� n  �� 4  ���]
�] 
cpar o  ���\�\ 0 i   o  ���[�[ "0 userinformation userInformation m  �� � ( d s A t t r T y p e N a t i v e : c o : Q  � r  �� n  �� !  4  ���Z"
�Z 
citm" m  ���Y�Y ! n  ��#$# 4  ���X%
�X 
cpar% o  ���W�W 0 i  $ o  ���V�V "0 userinformation userInformation o      �U�U 0 usercountry userCountry R      �T�S�R
�T .ascrerr ****      � ****�S  �R   k  �&& '(' r  ��)*) J  ��++ ,�Q, m  ��-- �..  �Q  * n     /0/ 1  ���P
�P 
txdl0 1  ���O
�O 
ascr( 1�N1 r  �232 c  �454 n  �676 7 �M89
�M 
cha 8 m  �L�L 9  ;  	
7 n  � :;: 4  � �K<
�K 
cpar< l ��=�J�I= [  ��>?> o  ���H�H 0 i  ? m  ���G�G �J  �I  ; o  ���F�F "0 userinformation userInformation5 m  �E
�E 
TEXT3 o      �D�D 0 usercountry userCountry�N  �_  �^   @A@ l �C�B�A�C  �B  �A  A BCB r  %DED J  FF G�@G m  HH �II  :  �@  E n     JKJ 1   $�?
�? 
txdlK 1   �>
�> 
ascrC LML Z  &�NO�=�<N C &2PQP n  &.RSR 4  ).�;T
�; 
cparT o  ,-�:�: 0 i  S o  &)�9�9 "0 userinformation userInformationQ m  .1UU �VV 2 d s A t t r T y p e N a t i v e : c o m p a n y :O Q  5|WXYW r  8IZ[Z n  8E\]\ 4  @E�8^
�8 
citm^ m  CD�7�7 ] n  8@_`_ 4  ;@�6a
�6 
cpara o  >?�5�5 0 i  ` o  8;�4�4 "0 userinformation userInformation[ o      �3�3 0 usercompany userCompanyX R      �2�1�0
�2 .ascrerr ****      � ****�1  �0  Y k  Q|bb cdc r  Q^efe J  QVgg h�/h m  QTii �jj  �/  f n     klk 1  Y]�.
�. 
txdll 1  VY�-
�- 
ascrd m�,m r  _|non c  _xpqp n  _trsr 7it�+tu
�+ 
cha t m  oq�*�* u  ;  rss n  _ivwv 4  bi�)x
�) 
cparx l ehy�(�'y [  ehz{z o  ef�&�& 0 i  { m  fg�%�% �(  �'  w o  _b�$�$ "0 userinformation userInformationq m  tw�#
�# 
TEXTo o      �"�" 0 usercompany userCompany�,  �=  �<  M |}| l ���!� ��!  �   �  } ~~ r  ����� J  ���� ��� m  ���� ���  :  �  � n     ��� 1  ���
� 
txdl� 1  ���
� 
ascr ��� Z  ������� C ����� n  ����� 4  ����
� 
cpar� o  ���� 0 i  � o  ���� "0 userinformation userInformation� m  ���� ��� 8 d s A t t r T y p e N a t i v e : d e p a r t m e n t :� Q  ������ r  ����� n  ����� 4  ����
� 
citm� m  ���� � n  ����� 4  ����
� 
cpar� o  ���� 0 i  � o  ���� "0 userinformation userInformation� o      ��  0 userdepartment userDepartment� R      ���
� .ascrerr ****      � ****�  �  � k  ���� ��� r  ����� J  ���� ��� m  ���� ���  �  � n     ��� 1  ���
� 
txdl� 1  ���
� 
ascr� ��
� r  ����� c  ����� n  ����� 7���	��
�	 
cha � m  ���� �  ;  ��� n  ����� 4  ����
� 
cpar� l ������ [  ����� o  ���� 0 i  � m  ���� �  �  � o  ���� "0 userinformation userInformation� m  ���
� 
TEXT� o      � �   0 userdepartment userDepartment�
  �  �  � ��� l ����������  ��  ��  � ��� r  ����� J  ���� ���� m  ���� ���  :  ��  � n     ��� 1  ����
�� 
txdl� 1  ����
�� 
ascr� ��� Z  �R������� C ���� n  � ��� 4  � ���
�� 
cpar� o  ������ 0 i  � o  ������ "0 userinformation userInformation� m   �� ��� X d s A t t r T y p e N a t i v e : p h y s i c a l D e l i v e r y O f f i c e N a m e :� Q  N���� r  
��� n  
��� 4  ���
�� 
citm� m  ���� � n  
��� 4  ���
�� 
cpar� o  ���� 0 i  � o  
���� "0 userinformation userInformation� o      ���� 0 
useroffice 
userOffice� R      ������
�� .ascrerr ****      � ****��  ��  � k  #N�� ��� r  #0��� J  #(�� ���� m  #&�� ���  ��  � n     ��� 1  +/��
�� 
txdl� 1  (+��
�� 
ascr� ���� r  1N��� c  1J��� n  1F��� 7;F����
�� 
cha � m  AC���� �  ;  DE� n  1;��� 4  4;���
�� 
cpar� l 7:������ [  7:��� o  78���� 0 i  � m  89���� ��  ��  � o  14���� "0 userinformation userInformation� m  FI��
�� 
TEXT� o      ���� 0 
useroffice 
userOffice��  ��  ��  � ��� l SS��������  ��  ��  � ��� r  S`��� J  SX�� ���� m  SV�� ���  :  ��  � n     ��� 1  [_��
�� 
txdl� 1  X[��
�� 
ascr�    Z  a����� C am n  ai 4  di��
�� 
cpar o  gh���� 0 i   o  ad���� "0 userinformation userInformation m  il		 �

 @ d s A t t r T y p e N a t i v e : s A M A c c o u n t N a m e : Q  p� r  s� n  s� 4  {���
�� 
citm m  ~����  n  s{ 4  v{��
�� 
cpar o  yz���� 0 i   o  sv���� "0 userinformation userInformation o      ���� 0 usershortname userShortName R      ������
�� .ascrerr ****      � ****��  ��   k  ��  r  �� J  �� �� m  �� �  ��   n       1  ����
�� 
txdl  1  ����
�� 
ascr !��! r  ��"#" c  ��$%$ n  ��&'& 7����()
�� 
cha ( m  ������ )  ;  ��' n  ��*+* 4  ����,
�� 
cpar, l ��-����- [  ��./. o  ������ 0 i  / m  ������ ��  ��  + o  ������ "0 userinformation userInformation% m  ����
�� 
TEXT# o      ���� 0 usershortname userShortName��  ��  ��   010 l ����������  ��  ��  1 232 r  ��454 J  ��66 7��7 m  ��88 �99  :  ��  5 n     :;: 1  ����
�� 
txdl; 1  ����
�� 
ascr3 <=< Z  �$>?����> C ��@A@ n  ��BCB 4  ����D
�� 
cparD o  ������ 0 i  C o  ������ "0 userinformation userInformationA m  ��EE �FF : d s A t t r T y p e N a t i v e : w W W H o m e P a g e :? Q  � GHIG r  ��JKJ n  ��LML 4  ����N
�� 
citmN m  ������ M n  ��OPO 4  ����Q
�� 
cparQ o  ������ 0 i  P o  ������ "0 userinformation userInformationK o      ���� 0 userwebpage userWebPageH R      ������
�� .ascrerr ****      � ****��  ��  I k  � RR STS r  �UVU J  ��WW X��X m  ��YY �ZZ  ��  V n     [\[ 1  ���
�� 
txdl\ 1  ����
�� 
ascrT ]��] r   ^_^ c  `a` n  bcb 7��de
�� 
cha d m  ���� e  ;  c n  fgf 4  ��h
�� 
cparh l 	i����i [  	jkj o  	
���� 0 i  k m  
���� ��  ��  g o  ���� "0 userinformation userInformationa m  ��
�� 
TEXT_ o      ���� 0 userwebpage userWebPage��  ��  ��  = lml l %%��������  ��  ��  m non r  %2pqp J  %*rr s��s m  %(tt �uu  :  ��  q n     vwv 1  -1��
�� 
txdlw 1  *-��
�� 
ascro xyx Z  3�z{����z C 3?|}| n  3;~~ 4  6;���
�� 
cpar� o  9:���� 0 i   o  36���� "0 userinformation userInformation} m  ;>�� ��� 
 C i t y :{ Q  B����� r  EV��� n  ER��� 4  MR���
�� 
citm� m  PQ���� � n  EM��� 4  HM���
�� 
cpar� o  KL���� 0 i  � o  EH���� "0 userinformation userInformation� o      ���� 0 usercity userCity� R      ������
�� .ascrerr ****      � ****��  ��  � k  ^��� ��� r  ^k��� J  ^c�� ���� m  ^a�� ���  ��  � n     ��� 1  fj��
�� 
txdl� 1  cf��
�� 
ascr� ���� r  l���� c  l���� n  l���� 7v�����
�� 
cha � m  |~���� �  ;  �� n  lv��� 4  ov��
� 
cpar� l ru��~�}� [  ru��� o  rs�|�| 0 i  � m  st�{�{ �~  �}  � o  lo�z�z "0 userinformation userInformation� m  ���y
�y 
TEXT� o      �x�x 0 usercity userCity��  ��  ��  y ��� l ���w�v�u�w  �v  �u  � ��� r  ����� J  ���� ��t� m  ���� ���  :  �t  � n     ��� 1  ���s
�s 
txdl� 1  ���r
�r 
ascr� ��� Z  �����q�p� C ����� n  ����� 4  ���o�
�o 
cpar� o  ���n�n 0 i  � o  ���m�m "0 userinformation userInformation� m  ���� ���  F A X N u m b e r :� Q  ������ r  ����� n  ����� 4  ���l�
�l 
citm� m  ���k�k � n  ����� 4  ���j�
�j 
cpar� o  ���i�i 0 i  � o  ���h�h "0 userinformation userInformation� o      �g�g 0 userfax userFax� R      �f�e�d
�f .ascrerr ****      � ****�e  �d  � k  ���� ��� r  ����� J  ���� ��c� m  ���� ���  �c  � n     ��� 1  ���b
�b 
txdl� 1  ���a
�a 
ascr� ��`� r  ����� c  ����� n  ����� 7���_��
�_ 
cha � m  ���^�^ �  ;  ��� n  ����� 4  ���]�
�] 
cpar� l ����\�[� [  ����� o  ���Z�Z 0 i  � m  ���Y�Y �\  �[  � o  ���X�X "0 userinformation userInformation� m  ���W
�W 
TEXT� o      �V�V 0 userfax userFax�`  �q  �p  � ��� l ���U�T�S�U  �T  �S  � ��� r  ���� J  ���� ��R� m  ���� ���  :  �R  � n     ��� 1  ��Q
�Q 
txdl� 1  ���P
�P 
ascr� ��� Z  _���O�N� C ��� n  ��� 4  �M�
�M 
cpar� o  �L�L 0 i  � o  �K�K "0 userinformation userInformation� m  �� ���  F i r s t N a m e :� Q  [���� r  (��� n  $   4  $�J
�J 
citm m  "#�I�I  n   4  �H
�H 
cpar o  �G�G 0 i   o  �F�F "0 userinformation userInformation� o      �E�E 0 userfirstname userFirstName� R      �D�C�B
�D .ascrerr ****      � ****�C  �B  � k  0[  r  0=	
	 J  05 �A m  03 �  �A  
 n      1  8<�@
�@ 
txdl 1  58�?
�? 
ascr �> r  >[ c  >W n  >S 7HS�=
�= 
cha  m  NP�<�<   ;  QR n  >H 4  AH�;
�; 
cpar l DG�:�9 [  DG o  DE�8�8 0 i   m  EF�7�7 �:  �9   o  >A�6�6 "0 userinformation userInformation m  SV�5
�5 
TEXT o      �4�4 0 userfirstname userFirstName�>  �O  �N  �  !  l ``�3�2�1�3  �2  �1  ! "#" r  `m$%$ J  `e&& '�0' m  `c(( �))  :  �0  % n     *+* 1  hl�/
�/ 
txdl+ 1  eh�.
�. 
ascr# ,-, Z  n�./�-�,. C nz010 n  nv232 4  qv�+4
�+ 
cpar4 o  tu�*�* 0 i  3 o  nq�)�) "0 userinformation userInformation1 m  vy55 �66  J o b T i t l e :/ Q  }�7897 r  ��:;: n  ��<=< 4  ���(>
�( 
citm> m  ���'�' = n  ��?@? 4  ���&A
�& 
cparA o  ���%�% 0 i  @ o  ���$�$ "0 userinformation userInformation; o      �#�# 0 	usertitle 	userTitle8 R      �"�!� 
�" .ascrerr ****      � ****�!  �   9 k  ��BB CDC r  ��EFE J  ��GG H�H m  ��II �JJ  �  F n     KLK 1  ���
� 
txdlL 1  ���
� 
ascrD M�M r  ��NON c  ��PQP n  ��RSR 7���TU
� 
cha T m  ���� U  ;  ��S n  ��VWV 4  ���X
� 
cparX l ��Y��Y [  ��Z[Z o  ���� 0 i  [ m  ���� �  �  W o  ���� "0 userinformation userInformationQ m  ���
� 
TEXTO o      �� 0 	usertitle 	userTitle�  �-  �,  - \]\ l ������  �  �  ] ^_^ r  ��`a` J  ��bb c�c m  ��dd �ee  :  �  a n     fgf 1  ���
� 
txdlg 1  ���
� 
ascr_ hih Z  �1jk��
j C ��lml n  ��non 4  ���	p
�	 
cparp o  ���� 0 i  o o  ���� "0 userinformation userInformationm m  ��qq �rr  L a s t N a m e :k Q  �-stus r  ��vwv n  ��xyx 4  ���z
� 
citmz m  ���� y n  ��{|{ 4  ���}
� 
cpar} o  ���� 0 i  | o  ���� "0 userinformation userInformationw o      �� 0 userlastname userLastNamet R      � ����
�  .ascrerr ****      � ****��  ��  u k  -~~ � r  ��� J  �� ���� m  �� ���  ��  � n     ��� 1  
��
�� 
txdl� 1  
��
�� 
ascr� ���� r  -��� c  )��� n  %��� 7%����
�� 
cha � m   "���� �  ;  #$� n  ��� 4  ���
�� 
cpar� l ������ [  ��� o  ���� 0 i  � m  ���� ��  ��  � o  ���� "0 userinformation userInformation� m  %(��
�� 
TEXT� o      ���� 0 userlastname userLastName��  �  �
  i ��� l 22��������  ��  ��  � ��� r  2?��� J  27�� ���� m  25�� ���  :  ��  � n     ��� 1  :>��
�� 
txdl� 1  7:��
�� 
ascr� ��� Z  @�������� C @L��� n  @H��� 4  CH���
�� 
cpar� o  FG���� 0 i  � o  @C���� "0 userinformation userInformation� m  HK�� ���  M o b i l e N u m b e r :� Q  O����� r  Rc��� n  R_��� 4  Z_���
�� 
citm� m  ]^���� � n  RZ��� 4  UZ���
�� 
cpar� o  XY���� 0 i  � o  RU���� "0 userinformation userInformation� o      ���� 0 
usermobile 
userMobile� R      ������
�� .ascrerr ****      � ****��  ��  � k  k��� ��� r  kx��� J  kp�� ���� m  kn�� ���  ��  � n     ��� 1  sw��
�� 
txdl� 1  ps��
�� 
ascr� ���� r  y���� c  y���� n  y���� 7������
�� 
cha � m  ������ �  ;  ��� n  y���� 4  |����
�� 
cpar� l ������� [  ���� o  ����� 0 i  � m  ������ ��  ��  � o  y|���� "0 userinformation userInformation� m  ����
�� 
TEXT� o      ���� 0 
usermobile 
userMobile��  ��  ��  � ��� l ����������  ��  ��  � ��� r  ����� J  ���� ���� m  ���� ���  :  ��  � n     ��� 1  ����
�� 
txdl� 1  ����
�� 
ascr� ��� Z  �	������� C ����� n  ����� 4  �����
�� 
cpar� o  ������ 0 i  � o  ������ "0 userinformation userInformation� m  ���� ���  P h o n e N u m b e r :� Q  ������ r  ����� n  ����� 4  �����
�� 
citm� m  ������ � n  ����� 4  �����
�� 
cpar� o  ������ 0 i  � o  ������ "0 userinformation userInformation� o      ���� 0 userworkphone userWorkPhone� R      ������
�� .ascrerr ****      � ****��  ��  � k  ���� ��� r  ����� J  ���� ���� m  ���� ���  ��  � n     � � 1  ����
�� 
txdl  1  ����
�� 
ascr� �� r  �� c  �� n  �� 7����	
�� 
cha  m  ������ 	  ;  �� n  ��

 4  ����
�� 
cpar l ������ [  �� o  ������ 0 i   m  ������ ��  ��   o  ������ "0 userinformation userInformation m  ����
�� 
TEXT o      ���� 0 userworkphone userWorkPhone��  ��  ��  �  l 		��������  ��  ��    r  		 J  			 �� m  		 �  :  ��   n      1  		��
�� 
txdl 1  			��
�� 
ascr  Z  		l���� C 		 !  n  		"#" 4  		��$
�� 
cpar$ o  		���� 0 i  # o  		���� "0 userinformation userInformation! m  		%% �&&  P o s t a l C o d e : Q  	!	h'()' r  	$	5*+* n  	$	1,-, 4  	,	1��.
�� 
citm. m  	/	0���� - n  	$	,/0/ 4  	'	,��1
�� 
cpar1 o  	*	+���� 0 i  0 o  	$	'���� "0 userinformation userInformation+ o      ����  0 userpostalcode userPostalCode( R      ������
�� .ascrerr ****      � ****��  ��  ) k  	=	h22 343 r  	=	J565 J  	=	B77 8��8 m  	=	@99 �::  ��  6 n     ;<; 1  	E	I��
�� 
txdl< 1  	B	E��
�� 
ascr4 =��= r  	K	h>?> c  	K	d@A@ n  	K	`BCB 7	U	`��DE
�� 
cha D m  	[	]���� E  ;  	^	_C n  	K	UFGF 4  	N	U��H
�� 
cparH l 	Q	TI����I [  	Q	TJKJ o  	Q	R���� 0 i  K m  	R	S���� ��  ��  G o  	K	N���� "0 userinformation userInformationA m  	`	c��
�� 
TEXT? o      ����  0 userpostalcode userPostalCode��  ��  ��   LML l 	m	m��������  ��  ��  M NON r  	m	zPQP J  	m	rRR S��S m  	m	pTT �UU  :  ��  Q n     VWV 1  	u	y��
�� 
txdlW 1  	r	u��
�� 
ascrO XYX Z  	{	�Z[����Z C 	{	�\]\ n  	{	�^_^ 4  	~	���`
�� 
cpar` o  	�	����� 0 i  _ o  	{	~�� "0 userinformation userInformation] m  	�	�aa �bb  R e a l N a m e :[ Q  	�	�cdec r  	�	�fgf n  	�	�hih 4  	�	��~j
�~ 
citmj m  	�	��}�} i n  	�	�klk 4  	�	��|m
�| 
cparm o  	�	��{�{ 0 i  l o  	�	��z�z "0 userinformation userInformationg o      �y�y 0 userfullname userFullNamed R      �x�w�v
�x .ascrerr ****      � ****�w  �v  e k  	�	�nn opo r  	�	�qrq J  	�	�ss t�ut m  	�	�uu �vv  �u  r n     wxw 1  	�	��t
�t 
txdlx 1  	�	��s
�s 
ascrp y�ry r  	�	�z{z c  	�	�|}| n  	�	�~~ 7	�	��q��
�q 
cha � m  	�	��p�p �  ;  	�	� n  	�	���� 4  	�	��o�
�o 
cpar� l 	�	���n�m� [  	�	���� o  	�	��l�l 0 i  � m  	�	��k�k �n  �m  � o  	�	��j�j "0 userinformation userInformation} m  	�	��i
�i 
TEXT{ o      �h�h 0 userfullname userFullName�r  ��  ��  Y ��� l 	�	��g�f�e�g  �f  �e  � ��� r  	�	���� J  	�	��� ��d� m  	�	��� ���  :  �d  � n     ��� 1  	�	��c
�c 
txdl� 1  	�	��b
�b 
ascr� ��� Z  	�
>���a�`� C 	�	���� n  	�	���� 4  	�	��_�
�_ 
cpar� o  	�	��^�^ 0 i  � o  	�	��]�] "0 userinformation userInformation� m  	�	��� ���  S t a t e :� Q  	�
:���� r  	�
��� n  	�
��� 4  	�
�\�
�\ 
citm� m  

�[�[ � n  	�	���� 4  	�	��Z�
�Z 
cpar� o  	�	��Y�Y 0 i  � o  	�	��X�X "0 userinformation userInformation� o      �W�W 0 	userstate 	userState� R      �V�U�T
�V .ascrerr ****      � ****�U  �T  � k  

:�� ��� r  

��� J  

�� ��S� m  

�� ���  �S  � n     ��� 1  

�R
�R 
txdl� 1  

�Q
�Q 
ascr� ��P� r  

:��� c  

6��� n  

2��� 7
'
2�O��
�O 
cha � m  
-
/�N�N �  ;  
0
1� n  

'��� 4  
 
'�M�
�M 
cpar� l 
#
&��L�K� [  
#
&��� o  
#
$�J�J 0 i  � m  
$
%�I�I �L  �K  � o  

 �H�H "0 userinformation userInformation� m  
2
5�G
�G 
TEXT� o      �F�F 0 	userstate 	userState�P  �a  �`  � ��� l 
?
?�E�D�C�E  �D  �C  � ��� r  
?
L��� J  
?
D�� ��B� m  
?
B�� ���  :  �B  � n     ��� 1  
G
K�A
�A 
txdl� 1  
D
G�@
�@ 
ascr� ��� Z  
M
����?�>� C 
M
Y��� n  
M
U��� 4  
P
U�=�
�= 
cpar� o  
S
T�<�< 0 i  � o  
M
P�;�; "0 userinformation userInformation� m  
U
X�� ���  S t r e e t :� Q  
\
����� r  
_
p��� n  
_
l��� 4  
g
l�:�
�: 
citm� m  
j
k�9�9 � n  
_
g��� 4  
b
g�8�
�8 
cpar� o  
e
f�7�7 0 i  � o  
_
b�6�6 "0 userinformation userInformation� o      �5�5 0 
userstreet 
userStreet� R      �4�3�2
�4 .ascrerr ****      � ****�3  �2  � k  
x
��� ��� r  
x
���� J  
x
}�� ��1� m  
x
{�� ���  �1  � n     ��� 1  
�
��0
�0 
txdl� 1  
}
��/
�/ 
ascr� ��.� r  
�
���� c  
�
���� n  
�
���� 7
�
��-��
�- 
cha � m  
�
��,�, �  ;  
�
�� n  
�
���� 4  
�
��+�
�+ 
cpar� l 
�
���*�)� [  
�
���� o  
�
��(�( 0 i  � m  
�
��'�' �*  �)  � o  
�
��&�& "0 userinformation userInformation� m  
�
��%
�% 
TEXT� o      �$�$ 0 
userstreet 
userStreet�.  �?  �>  � 	 �#	  l 
�
��"�!� �"  �!  �   �#  �� 0 i  � m  56�� � I 6A�	�
� .corecnte****       ****	 n 6=			 2 9=�
� 
cpar	 o  69�� "0 userinformation userInformation�  ��  � 			 l 
�
�����  �  �  	 			 r  
�
�				 J  
�
�	
	
 			 m  
�
�		 �		  ; K e r b e r o s v 5 ; ;	 	�	 m  
�
�		 �		  ;�  		 n     			 1  
�
��
� 
txdl	 1  
�
��
� 
ascr	 			 l 
�
�����  �  �  	 			 Q  
�
�		�	 r  
�
�			 n  
�
�			 4  
�
��	
� 
citm	 m  
�
��� 	 o  
�
��� "0 userinformation userInformation	 o      �� &0 userkerberosrealm userKerberosRealm	 R      ���

� .ascrerr ****      � ****�  �
  �  	 		 	 l 
�
��	���	  �  �  	  	!	"	! r  
�
�	#	$	# J  
�
�	%	% 	&�	& m  
�
�	'	' �	(	(  �  	$ n     	)	*	) 1  
�
��
� 
txdl	* 1  
�
��
� 
ascr	" 	+	,	+ l 
�
�����  �  �  	, 	-	.	- Z  
�8	/	0� ��	/ = 
�
�	1	2	1 o  
�
����� 0 emailaddress emailAddress	2 m  
�
�	3	3 �	4	4  	0 k  
�4	5	5 	6	7	6 l 
�
���������  ��  ��  	7 	8	9	8 l 
�
���	:	;��  	:   something went wrong   	; �	<	< *   s o m e t h i n g   w e n t   w r o n g	9 	=	>	= l 
�
���������  ��  ��  	> 	?	@	? I 
� ��	A	B
�� .sysodlogaskr        TEXT	A b  
�
�	C	D	C b  
�
�	E	F	E b  
�
�	G	H	G o  
�
����� 0 errormessage errorMessage	H o  
�
���
�� 
ret 	F o  
�
���
�� 
ret 	D m  
�
�	I	I �	J	J h U n a b l e   t o   r e a d   e m a i l   a d d r e s s   f r o m   n e t w o r k   d i r e c t o r y .	B ��	K	L
�� 
disp	K m  ��
�� stic    	L ��	M	N
�� 
btns	M J  	O	O 	P��	P m  
	Q	Q �	R	R  O K��  	N ��	S	T
�� 
dflt	S J  	U	U 	V��	V m  	W	W �	X	X  O K��  	T ��	Y��
�� 
appr	Y m  	Z	Z �	[	[ , O u t l o o k   E x c h a n g e   S e t u p��  	@ 	\	]	\ I  !)��	^���� 0 writelog writeLog	^ 	_��	_ m  "%	`	` �	a	a v E R R O R :   U n a b l e   t o   r e a d   e m a i l   a d d r e s s   f r o m   n e t w o r k   d i r e c t o r y .��  ��  	] 	b��	b R  *4����	c
�� .ascrerr ****      � ****��  	c ��	d��
�� 
errn	d m  .1��������  ��  �   ��  	. 	e��	e l 99��������  ��  ��  ��  � 	f	g	f F  =R	h	i	h = =D	j	k	j o  =B���� 0 emailformat emailFormat	k m  BC���� 	i = GN	l	m	l o  GL���� 0 displayname displayName	m m  LM���� 	g 	n	o	n k  U�	p	p 	q	r	q l UU��������  ��  ��  	r 	s	t	s l UU��	u	v��  	u P J Pull user information from the account settings of the local user account   	v �	w	w �   P u l l   u s e r   i n f o r m a t i o n   f r o m   t h e   a c c o u n t   s e t t i n g s   o f   t h e   l o c a l   u s e r   a c c o u n t	t 	x	y	x l UU��������  ��  ��  	y 	z	{	z r  Ub	|	}	| n  U^	~		~ 1  Z^��
�� 
sisn	 l UZ	�����	� I UZ������
�� .sysosigtsirr   ��� null��  ��  ��  ��  	} o      ���� 0 usershortname userShortName	{ 	�	�	� r  cp	�	�	� n  cl	�	�	� 1  hl��
�� 
siln	� l ch	�����	� I ch������
�� .sysosigtsirr   ��� null��  ��  ��  ��  	� o      ���� 0 userfullname userFullName	� 	�	�	� l qq��������  ��  ��  	� 	�	�	� l qq��	�	���  	� D > first.last@domain.com and full name displays as "Last, First"   	� �	�	� |   f i r s t . l a s t @ d o m a i n . c o m   a n d   f u l l   n a m e   d i s p l a y s   a s   " L a s t ,   F i r s t "	� 	�	�	� l qq��������  ��  ��  	� 	�	�	� r  q|	�	�	� m  qt	�	� �	�	�  ,  	� n     	�	�	� 1  w{��
�� 
txdl	� 1  tw��
�� 
ascr	� 	�	�	� r  }�	�	�	� n  }�	�	�	� 4 ����	�
�� 
citm	� m  ��������	� o  }����� 0 userfullname userFullName	� o      ���� 0 userfirstname userFirstName	� 	�	�	� r  ��	�	�	� n  ��	�	�	� 4  ����	�
�� 
cwor	� m  ������ 	� n  ��	�	�	� 4  ����	�
�� 
citm	� m  ������ 	� o  ������ 0 userfullname userFullName	� o      ���� 0 userlastname userLastName	� 	�	�	� r  ��	�	�	� m  ��	�	� �	�	�  	� n     	�	�	� 1  ����
�� 
txdl	� 1  ����
�� 
ascr	� 	�	�	� r  ��	�	�	� b  ��	�	�	� b  ��	�	�	� b  ��	�	�	� b  ��	�	�	� o  ������ 0 userfirstname userFirstName	� m  ��	�	� �	�	�  .	� o  ������ 0 userlastname userLastName	� m  ��	�	� �	�	�  @	� o  ������ 0 
domainname 
domainName	� o      ���� 0 emailaddress emailAddress	� 	���	� l ����������  ��  ��  ��  	o 	�	�	� F  ��	�	�	� = ��	�	�	� o  ������ 0 emailformat emailFormat	� m  ������ 	� = ��	�	�	� o  ������ 0 displayname displayName	� m  ������ 	� 	�	�	� k  �L	�	� 	�	�	� l ����������  ��  ��  	� 	�	�	� l ����	�	���  	� P J Pull user information from the account settings of the local user account   	� �	�	� �   P u l l   u s e r   i n f o r m a t i o n   f r o m   t h e   a c c o u n t   s e t t i n g s   o f   t h e   l o c a l   u s e r   a c c o u n t	� 	�	�	� l ����������  ��  ��  	� 	�	�	� r  ��	�	�	� n  ��	�	�	� 1  ����
�� 
sisn	� l ��	�����	� I ��������
�� .sysosigtsirr   ��� null��  ��  ��  ��  	� o      ���� 0 usershortname userShortName	� 	�	�	� r  ��	�	�	� n  ��	�	�	� 1  ����
�� 
siln	� l ��	�����	� I ��������
�� .sysosigtsirr   ��� null��  ��  ��  ��  	� o      ���� 0 userfullname userFullName	� 	�	�	� l ����������  ��  ��  	� 	�	�	� l ����	�	���  	� C = first.last@domain.com and full name displays as "First Last"   	� �	�	� z   f i r s t . l a s t @ d o m a i n . c o m   a n d   f u l l   n a m e   d i s p l a y s   a s   " F i r s t   L a s t "	� 	�	�	� l ����������  ��  ��  	� 	�	�	� r  �	�	�	� m  ��	�	� �	�	�   	� n     	�	�	� 1   ��
�� 
txdl	� 1  � ��
�� 
ascr	� 	�	�	� r  	�	�	� n  	�	�	� 4  ��	�
�� 
cwor	� m  ���� 	� n  	�	�	� 4  	��	�
�� 
citm	� m  ���� 	� o  	���� 0 userfullname userFullName	� o      ���� 0 userfirstname userFirstName	� 
 

  r  $


 n   


 4  �

� 
citm
 m  �~�~��
 o  �}�} 0 userfullname userFullName
 o      �|�| 0 userlastname userLastName
 


 r  %0
	


	 m  %(

 �

  

 n     


 1  +/�{
�{ 
txdl
 1  (+�z
�z 
ascr
 


 r  1J


 b  1F


 b  1@


 b  1<


 b  18


 o  14�y�y 0 userfirstname userFirstName
 m  47

 �

  .
 o  8;�x�x 0 userlastname userLastName
 m  <?

 �

  @
 o  @E�w�w 0 
domainname 
domainName
 o      �v�v 0 emailaddress emailAddress
 
�u
 l KK�t�s�r�t  �s  �r  �u  	� 
 
!
  F  Od
"
#
" = OV
$
%
$ o  OT�q�q 0 emailformat emailFormat
% m  TU�p�p 
# = Y`
&
'
& o  Y^�o�o 0 displayname displayName
' m  ^_�n�n 
! 
(
)
( k  g�
*
* 
+
,
+ l gg�m�l�k�m  �l  �k  
, 
-
.
- l gg�j
/
0�j  
/ P J Pull user information from the account settings of the local user account   
0 �
1
1 �   P u l l   u s e r   i n f o r m a t i o n   f r o m   t h e   a c c o u n t   s e t t i n g s   o f   t h e   l o c a l   u s e r   a c c o u n t
. 
2
3
2 l gg�i�h�g�i  �h  �g  
3 
4
5
4 r  gt
6
7
6 n  gp
8
9
8 1  lp�f
�f 
sisn
9 l gl
:�e�d
: I gl�c�b�a
�c .sysosigtsirr   ��� null�b  �a  �e  �d  
7 o      �`�` 0 usershortname userShortName
5 
;
<
; r  u�
=
>
= n  u~
?
@
? 1  z~�_
�_ 
siln
@ l uz
A�^�]
A I uz�\�[�Z
�\ .sysosigtsirr   ��� null�[  �Z  �^  �]  
> o      �Y�Y 0 userfullname userFullName
< 
B
C
B l ���X�W�V�X  �W  �V  
C 
D
E
D l ���U
F
G�U  
F ? 9 first@domain.com and full name displays as "Last, First"   
G �
H
H r   f i r s t @ d o m a i n . c o m   a n d   f u l l   n a m e   d i s p l a y s   a s   " L a s t ,   F i r s t "
E 
I
J
I l ���T�S�R�T  �S  �R  
J 
K
L
K r  ��
M
N
M m  ��
O
O �
P
P  ,  
N n     
Q
R
Q 1  ���Q
�Q 
txdl
R 1  ���P
�P 
ascr
L 
S
T
S r  ��
U
V
U n  ��
W
X
W 4 ���O
Y
�O 
citm
Y m  ���N�N��
X o  ���M�M 0 userfullname userFullName
V o      �L�L 0 userfirstname userFirstName
T 
Z
[
Z r  ��
\
]
\ n  ��
^
_
^ 4  ���K
`
�K 
cwor
` m  ���J�J 
_ n  ��
a
b
a 4  ���I
c
�I 
citm
c m  ���H�H 
b o  ���G�G 0 userfullname userFullName
] o      �F�F 0 userlastname userLastName
[ 
d
e
d r  ��
f
g
f m  ��
h
h �
i
i  
g n     
j
k
j 1  ���E
�E 
txdl
k 1  ���D
�D 
ascr
e 
l
m
l r  ��
n
o
n b  ��
p
q
p b  ��
r
s
r o  ���C�C 0 userfirstname userFirstName
s m  ��
t
t �
u
u  @
q o  ���B�B 0 
domainname 
domainName
o o      �A�A 0 emailaddress emailAddress
m 
v�@
v l ���?�>�=�?  �>  �=  �@  
) 
w
x
w F  ��
y
z
y = ��
{
|
{ o  ���<�< 0 emailformat emailFormat
| m  ���;�; 
z = ��
}
~
} o  ���:�: 0 displayname displayName
~ m  ���9�9 
x 

�
 k  �N
�
� 
�
�
� l ���8�7�6�8  �7  �6  
� 
�
�
� l ���5
�
��5  
� P J Pull user information from the account settings of the local user account   
� �
�
� �   P u l l   u s e r   i n f o r m a t i o n   f r o m   t h e   a c c o u n t   s e t t i n g s   o f   t h e   l o c a l   u s e r   a c c o u n t
� 
�
�
� l ���4�3�2�4  �3  �2  
� 
�
�
� r  ��
�
�
� n  ��
�
�
� 1  ���1
�1 
sisn
� l ��
��0�/
� I ���.�-�,
�. .sysosigtsirr   ��� null�-  �,  �0  �/  
� o      �+�+ 0 usershortname userShortName
� 
�
�
� r  �
�
�
� n  ��
�
�
� 1  ���*
�* 
siln
� l ��
��)�(
� I ���'�&�%
�' .sysosigtsirr   ��� null�&  �%  �)  �(  
� o      �$�$ 0 userfullname userFullName
� 
�
�
� l �#�"�!�#  �"  �!  
� 
�
�
� l � 
�
��   
� = 7 first@domain.com if full name displays as "First Last"   
� �
�
� n   f i r s t @ d o m a i n . c o m   i f   f u l l   n a m e   d i s p l a y s   a s   " F i r s t   L a s t "
� 
�
�
� l ����  �  �  
� 
�
�
� r  
�
�
� m  
�
� �
�
�   
� n     
�
�
� 1  
�
� 
txdl
� 1  
�
� 
ascr
� 
�
�
� r  !
�
�
� n  
�
�
� 4  �
�
� 
cwor
� m  �� 
� n  
�
�
� 4  �
�
� 
citm
� m  �� 
� o  �� 0 userfullname userFullName
� o      �� 0 userfirstname userFirstName
� 
�
�
� r  ".
�
�
� n  "*
�
�
� 4 %*�
�
� 
citm
� m  ()����
� o  "%�� 0 userfullname userFullName
� o      �� 0 userlastname userLastName
� 
�
�
� r  /:
�
�
� m  /2
�
� �
�
�  
� n     
�
�
� 1  59�
� 
txdl
� 1  25�
� 
ascr
� 
�
�
� r  ;L
�
�
� b  ;H
�
�
� b  ;B
�
�
� o  ;>�� 0 userfirstname userFirstName
� m  >A
�
� �
�
�  @
� o  BG�� 0 
domainname 
domainName
� o      �� 0 emailaddress emailAddress
� 
��
� l MM�
�	��
  �	  �  �  
� 
�
�
� F  Qf
�
�
� = QX
�
�
� o  QV�� 0 emailformat emailFormat
� m  VW�� 
� = [b
�
�
� o  [`�� 0 displayname displayName
� m  `a�� 
� 
�
�
� k  i�
�
� 
�
�
� l ii����  �  �  
� 
�
�
� l ii� 
�
��   
� P J Pull user information from the account settings of the local user account   
� �
�
� �   P u l l   u s e r   i n f o r m a t i o n   f r o m   t h e   a c c o u n t   s e t t i n g s   o f   t h e   l o c a l   u s e r   a c c o u n t
� 
�
�
� l ii��������  ��  ��  
� 
�
�
� r  iv
�
�
� n  ir
�
�
� 1  nr��
�� 
sisn
� l in
�����
� I in������
�� .sysosigtsirr   ��� null��  ��  ��  ��  
� o      ���� 0 usershortname userShortName
� 
�
�
� r  w�
�
�
� n  w�
�
�
� 1  |���
�� 
siln
� l w|
�����
� I w|������
�� .sysosigtsirr   ��� null��  ��  ��  ��  
� o      ���� 0 userfullname userFullName
� 
�
�
� l ����������  ��  ��  
� 
�
�
� l ����
�
���  
� ? 9 flast@domain.com and full name displays as "Last, First"   
� �
�
� r   f l a s t @ d o m a i n . c o m   a n d   f u l l   n a m e   d i s p l a y s   a s   " L a s t ,   F i r s t "
� 
�
�
� l ����������  ��  ��  
� 
�
�
� r  ��
�
�
� m  ��
�
� �
�
�  ,  
� n     
� 
� 1  ����
�� 
txdl  1  ����
�� 
ascr
�  r  �� n  �� 4 ����
�� 
citm m  �������� o  ������ 0 userfullname userFullName o      ���� 0 userfirstname userFirstName 	 r  ��

 n  �� 4  ����
�� 
cwor m  ������  n  �� 4  ����
�� 
citm m  ������  o  ������ 0 userfullname userFullName o      ���� 0 userlastname userLastName	  r  �� m  �� �   n      1  ����
�� 
txdl 1  ����
�� 
ascr  r  �� b  �� b  �� !  b  ��"#" l ��$����$ n  ��%&% 4  ����'
�� 
cha ' m  ������ & o  ������ 0 userfirstname userFirstName��  ��  # o  ������ 0 userlastname userLastName! m  ��(( �))  @ o  ������ 0 
domainname 
domainName o      ���� 0 emailaddress emailAddress *��* l ����������  ��  ��  ��  
� +,+ F  ��-.- = ��/0/ o  ������ 0 emailformat emailFormat0 m  ������ . = ��121 o  ������ 0 displayname displayName2 m  ������ , 343 k  �b55 676 l ����������  ��  ��  7 898 l ����:;��  : P J Pull user information from the account settings of the local user account   ; �<< �   P u l l   u s e r   i n f o r m a t i o n   f r o m   t h e   a c c o u n t   s e t t i n g s   o f   t h e   l o c a l   u s e r   a c c o u n t9 =>= l ����������  ��  ��  > ?@? r  � ABA n  ��CDC 1  ����
�� 
sisnD l ��E����E I ��������
�� .sysosigtsirr   ��� null��  ��  ��  ��  B o      ���� 0 usershortname userShortName@ FGF r  HIH n  
JKJ 1  
��
�� 
silnK l L����L I ������
�� .sysosigtsirr   ��� null��  ��  ��  ��  I o      ���� 0 userfullname userFullNameG MNM l ��������  ��  ��  N OPO l ��QR��  Q > 8 flast@domain.com and full name displays as "First Last"   R �SS p   f l a s t @ d o m a i n . c o m   a n d   f u l l   n a m e   d i s p l a y s   a s   " F i r s t   L a s t "P TUT l ��������  ��  ��  U VWV r  XYX m  ZZ �[[   Y n     \]\ 1  ��
�� 
txdl] 1  ��
�� 
ascrW ^_^ r  ,`a` n  (bcb 4  #(��d
�� 
cword m  &'���� c n  #efe 4  #��g
�� 
citmg m  !"���� f o  ���� 0 userfullname userFullNamea o      ���� 0 userfirstname userFirstName_ hih r  -9jkj n  -5lml 4 05��n
�� 
citmn m  34������m o  -0���� 0 userfullname userFullNamek o      ���� 0 userlastname userLastNamei opo r  :Eqrq m  :=ss �tt  r n     uvu 1  @D��
�� 
txdlv 1  =@��
�� 
ascrp wxw r  F`yzy l F\{����{ b  F\|}| b  FV~~ b  FR��� n  FN��� 4  IN���
�� 
cha � m  LM���� � o  FI���� 0 userfirstname userFirstName� o  NQ���� 0 userlastname userLastName m  RU�� ���  @} o  V[���� 0 
domainname 
domainName��  ��  z o      ���� 0 emailaddress emailAddressx ���� l aa��������  ��  ��  ��  4 ��� F  e|��� = en��� o  ej���� 0 emailformat emailFormat� m  jm���� � = qx��� o  qv���� 0 displayname displayName� m  vw���� � ��� k  ��� ��� l ��������  ��  ��  � ��� l ������  � P J Pull user information from the account settings of the local user account   � ��� �   P u l l   u s e r   i n f o r m a t i o n   f r o m   t h e   a c c o u n t   s e t t i n g s   o f   t h e   l o c a l   u s e r   a c c o u n t� ��� l ��������  ��  ��  � ��� r  ���� n  ���� 1  ����
�� 
sisn� l ������� I �������
�� .sysosigtsirr   ��� null��  ��  ��  ��  � o      ���� 0 usershortname userShortName� ��� r  ����� n  ����� 1  ����
�� 
siln� l ������� I ���~�}�|
�~ .sysosigtsirr   ��� null�}  �|  ��  �  � o      �{�{ 0 userfullname userFullName� ��� l ���z�y�x�z  �y  �x  � ��� l ���w���w  � C = shortName@domain.com and full name displays as "Last, First"   � ��� z   s h o r t N a m e @ d o m a i n . c o m   a n d   f u l l   n a m e   d i s p l a y s   a s   " L a s t ,   F i r s t "� ��� l ���v�u�t�v  �u  �t  � ��� r  ����� m  ���� ���  ,  � n     ��� 1  ���s
�s 
txdl� 1  ���r
�r 
ascr� ��� r  ����� n  ����� 4 ���q�
�q 
citm� m  ���p�p��� o  ���o�o 0 userfullname userFullName� o      �n�n 0 userfirstname userFirstName� ��� r  ����� n  ����� 4  ���m�
�m 
cwor� m  ���l�l � n  ����� 4  ���k�
�k 
citm� m  ���j�j � o  ���i�i 0 userfullname userFullName� o      �h�h 0 userlastname userLastName� ��� r  ����� m  ���� ���  � n     ��� 1  ���g
�g 
txdl� 1  ���f
�f 
ascr� ��� r  ����� b  ����� b  ����� o  ���e�e 0 usershortname userShortName� m  ���� ���  @� o  ���d�d 0 
domainname 
domainName� o      �c�c 0 emailaddress emailAddress� ��b� l ���a�`�_�a  �`  �_  �b  � ��� F  ����� = ����� o  ���^�^ 0 emailformat emailFormat� m  ���]�] � = ����� o  ���\�\ 0 displayname displayName� m  ���[�[ � ��Z� k  h�� ��� l �Y�X�W�Y  �X  �W  � ��� l �V���V  � P J Pull user information from the account settings of the local user account   � ��� �   P u l l   u s e r   i n f o r m a t i o n   f r o m   t h e   a c c o u n t   s e t t i n g s   o f   t h e   l o c a l   u s e r   a c c o u n t� ��� l �U�T�S�U  �T  �S  � ��� r  ��� n  ��� 1  �R
�R 
sisn� l ��Q�P� I �O�N�M
�O .sysosigtsirr   ��� null�N  �M  �Q  �P  � o      �L�L 0 usershortname userShortName� ��� r  ��� n  ��� 1  �K
�K 
siln� l ��J�I� I �H�G�F
�H .sysosigtsirr   ��� null�G  �F  �J  �I  � o      �E�E 0 userfullname userFullName�    l �D�C�B�D  �C  �B    l �A�A   B < shortName@domain.com and full name displays as "First Last"    � x   s h o r t N a m e @ d o m a i n . c o m   a n d   f u l l   n a m e   d i s p l a y s   a s   " F i r s t   L a s t "  l �@�?�>�@  �?  �>   	
	 r  ) m  ! �    n      1  $(�=
�= 
txdl 1  !$�<
�< 
ascr
  r  *; n  *7 4  27�;
�; 
cwor m  56�:�:  n  *2 4  -2�9
�9 
citm m  01�8�8  o  *-�7�7 0 userfullname userFullName o      �6�6 0 userfirstname userFirstName  r  <H n  <D  4 ?D�5!
�5 
citm! m  BC�4�4��  o  <?�3�3 0 userfullname userFullName o      �2�2 0 userlastname userLastName "#" r  IT$%$ m  IL&& �''  % n     ()( 1  OS�1
�1 
txdl) 1  LO�0
�0 
ascr# *+* r  Uf,-, b  Ub./. b  U\010 o  UX�/�/ 0 usershortname userShortName1 m  X[22 �33  @/ o  \a�.�. 0 
domainname 
domainName- o      �-�- 0 emailaddress emailAddress+ 4�,4 l gg�+�*�)�+  �*  �)  �,  �Z  � k  k�55 676 l kk�(�'�&�(  �'  �&  7 898 l kk�%:;�%  :   something went wrong   ; �<< *   s o m e t h i n g   w e n t   w r o n g9 =>= l kk�$�#�"�$  �#  �"  > ?@? I k��!AB
�! .sysodlogaskr        TEXTA b  kxCDC b  ktEFE b  krGHG o  kp� �  0 errormessage errorMessageH o  pq�
� 
ret F o  rs�
� 
ret D m  twII �JJ x U n a b l e   t o   p a r s e   a c c o u n t   i n f o r m a t i o n   f r o m   l o c a l   O S   X   a c c o u n t .B �KL
� 
dispK m  {~�
� stic    L �MN
� 
btnsM J  ��OO P�P m  ��QQ �RR  O K�  N �ST
� 
dfltS J  ��UU V�V m  ��WW �XX  O K�  T �Y�
� 
apprY m  ��ZZ �[[ , O u t l o o k   E x c h a n g e   S e t u p�  @ \]\ R  ����^
� .ascrerr ****      � ****�  ^ �_�
� 
errn_ m  �������  ] `�` l ������  �  �  �  �  �  � aba l     ���
�  �  �
  b cdc l     �	ef�	  e P J Substitute email address for username for mail systems such as Office 365   f �gg �   S u b s t i t u t e   e m a i l   a d d r e s s   f o r   u s e r n a m e   f o r   m a i l   s y s t e m s   s u c h   a s   O f f i c e   3 6 5d hih l     ����  �  �  i jkj l ��l��l Z  ��mn��m = ��opo o  ���� *0 useemailforusername useEmailForUsernamep m  ��� 
�  boovtruen r  ��qrq o  ������ 0 emailaddress emailAddressr o      ���� 0 usershortname userShortName�  �  �  �  k sts l     ��������  ��  ��  t uvu l     ��wx��  w 0 *------------------------------------------   x �yy T - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -v z{z l     ��|}��  | &   End collecting user information   } �~~ @   E n d   c o l l e c t i n g   u s e r   i n f o r m a t i o n{ � l     ������  � 0 *------------------------------------------   � ��� T - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -� ��� l     ��������  ��  ��  � ��� l     ������  � 0 *------------------------------------------   � ��� T - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -� ��� l     ������  � %  Begin logging user information   � ��� >   B e g i n   l o g g i n g   u s e r   i n f o r m a t i o n� ��� l     ������  � 0 *------------------------------------------   � ��� T - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -� ��� l     ��������  ��  ��  � ��� l �������� I  ��������� 0 writelog writeLog� ���� m  ���� ��� & U s e r   i n f o r m a t i o n . . .��  ��  ��  ��  � ��� l �������� I  ��������� 0 writelog writeLog� ���� b  ����� m  ���� ���  F i r s t   N a m e :  � o  ������ 0 userfirstname userFirstName��  ��  ��  ��  � ��� l �������� I  ��������� 0 writelog writeLog� ���� b  ����� m  ���� ���  L a s t   N a m e :  � o  ������ 0 userlastname userLastName��  ��  ��  ��  � ��� l �������� I  ��������� 0 writelog writeLog� ���� b  ����� m  ���� ���  E m a i l   A d d r e s s :  � o  ������ 0 emailaddress emailAddress��  ��  ��  ��  � ��� l �������� I  ��������� 0 writelog writeLog� ���� b  ����� m  ���� ���  D e p a r t m e n t :  � o  ������  0 userdepartment userDepartment��  ��  ��  ��  � ��� l ������� I  �������� 0 writelog writeLog� ���� b  ���� m  ���� ���  O f f i c e :  � o  ����� 0 
useroffice 
userOffice��  ��  ��  ��  � ��� l ������ I  ������� 0 writelog writeLog� ���� b  	��� m  	�� ���  C o m p a n y :  � o  ���� 0 usercompany userCompany��  ��  ��  ��  � ��� l !������ I  !������� 0 writelog writeLog� ���� b  ��� m  �� ���  W o r k   P h o n e :  � o  ���� 0 userworkphone userWorkPhone��  ��  ��  ��  � ��� l ".������ I  ".������� 0 writelog writeLog� ���� b  #*��� m  #&�� ���  M o b i l e   P h o n e :  � o  &)���� 0 
usermobile 
userMobile��  ��  ��  ��  � ��� l /;������ I  /;������� 0 writelog writeLog� ���� b  07��� m  03�� ��� 
 F A X :  � o  36���� 0 userfax userFax��  ��  ��  ��  � ��� l <H������ I  <H������� 0 writelog writeLog� ���� b  =D��� m  =@�� ���  T i t l e :  � o  @C���� 0 	usertitle 	userTitle��  ��  ��  ��  � ��� l IU������ I  IU������� 0 writelog writeLog� ���� b  JQ��� m  JM�� �    S t r e e t :  � o  MP���� 0 
userstreet 
userStreet��  ��  ��  ��  �  l Vb���� I  Vb������ 0 writelog writeLog �� b  W^ m  WZ �		  C i t y :   o  Z]���� 0 usercity userCity��  ��  ��  ��   

 l co���� I  co������ 0 writelog writeLog �� b  dk m  dg �  S t a t e :   o  gj���� 0 	userstate 	userState��  ��  ��  ��    l p|���� I  p|������ 0 writelog writeLog �� b  qx m  qt �  P o s t a l   C o d e :   o  tw����  0 userpostalcode userPostalCode��  ��  ��  ��    l }����� I  }������� 0 writelog writeLog  ��  b  ~�!"! m  ~�## �$$  C o u n t r y :  " o  ������ 0 usercountry userCountry��  ��  ��  ��   %&% l ��'����' I  ����(���� 0 writelog writeLog( )��) b  ��*+* m  ��,, �--  W e b   P a g e :  + o  ������ 0 userwebpage userWebPage��  ��  ��  ��  & ./. l ��0����0 I  ����1���� 0 writelog writeLog1 2��2 o  ����
�� 
ret ��  ��  ��  ��  / 343 l     ��������  ��  ��  4 565 l     ��78��  7 0 *------------------------------------------   8 �99 T - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -6 :;: l     �<=�  < #  End logging user information   = �>> :   E n d   l o g g i n g   u s e r   i n f o r m a t i o n; ?@? l     �~AB�~  A 0 *------------------------------------------   B �CC T - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -@ DED l     �}�|�{�}  �|  �{  E FGF l     �zHI�z  H 0 *------------------------------------------   I �JJ T - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -G KLK l     �yMN�y  M   Begin account setup   N �OO (   B e g i n   a c c o u n t   s e t u pL PQP l     �xRS�x  R 0 *------------------------------------------   S �TT T - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -Q UVU l     �w�v�u�w  �v  �u  V WXW l ��Y�t�sY O  ��Z[Z k  ��\\ ]^] I ���r�q�p
�r .miscactvnull��� ��� null�q  �p  ^ _`_ l ���o�n�m�o  �n  �m  ` aba Q  ��cdec k  ��ff ghg r  ��iji m  ���l
�l boovtruej 1  ���k
�k 
wkOfh k�jk n ��lml I  ���in�h�i 0 writelog writeLogn o�go m  ��pp �qq d S e t   M i c r o s o f t   O u t l o o k   t o   w o r k   o f f l i n e :   S u c c e s s f u l .�g  �h  m  f  ���j  d R      �f�e�d
�f .ascrerr ****      � ****�e  �d  e n ��rsr I  ���ct�b�c 0 writelog writeLogt u�au m  ��vv �ww \ S e t   M i c r o s o f t   O u t l o o k   t o   w o r k   o f f l i n e :   F a i l e d .�a  �b  s  f  ��b xyx l ���`�_�^�`  �_  �^  y z{z Q  �	|}~| k  �� ��� r  ����� o  ���]�] 0 unifiedinbox unifiedInbox� 1  ���\
�\ 
GrpF� ��[� n ����� I  ���Z��Y�Z 0 writelog writeLog� ��X� b  ����� b  ����� m  ���� ��� : S e t   G r o u p   S i m i l a r   F o l d e r s   t o  � o  ���W�W 0 unifiedinbox unifiedInbox� m  ���� ���  :   S u c c e s s f u l .�X  �Y  �  f  ���[  } R      �V�U�T
�V .ascrerr ****      � ****�U  �T  ~ n �	��� I  �	�S��R�S 0 writelog writeLog� ��Q� b  ���� b  ���� m  ���� ��� : S e t   G r o u p   S i m i l a r   F o l d e r s   t o  � o  � �P�P 0 unifiedinbox unifiedInbox� m  �� ���  :   F a i l e d .�Q  �R  �  f  ��{ ��� l 

�O�N�M�O  �N  �M  � ��� Q  
E���� k  +�� ��� r  ��� o  �L�L 20 hideonmycomputerfolders hideOnMyComputerFolders� 1  �K
�K 
hOMC� ��J� n +��� I  +�I��H�I 0 writelog writeLog� ��G� b  '��� b  #��� m  �� ��� F S e t   H i d e   O n   M y   C o m p u t e r   F o l d e r s   t o  � o  "�F�F 20 hideonmycomputerfolders hideOnMyComputerFolders� m  #&�� ���  :   S u c c e s s f u l .�G  �H  �  f  �J  � R      �E�D�C
�E .ascrerr ****      � ****�D  �C  � n 3E��� I  4E�B��A�B 0 writelog writeLog� ��@� b  4A��� b  4=��� m  47�� ��� F S e t   H i d e   O n   M y   C o m p u t e r   F o l d e r s   t o  � o  7<�?�? 20 hideonmycomputerfolders hideOnMyComputerFolders� m  =@�� ���  :   F a i l e d .�@  �A  �  f  34� ��� l FF�>�=�<�>  �=  �<  � ��� Z  F����;�:� = FM��� o  FK�9�9 (0 verifyemailaddress verifyEMailAddress� m  KL�8
�8 boovtrue� k  P��� ��� r  P���� I P|�7��
�7 .sysodlogaskr        TEXT� m  PS�� ��� X P l e a s e   v e r i f y   y o u r   e m a i l   a d d r e s s   i s   c o r r e c t .� �6��
�6 
dtxt� o  VY�5�5 0 emailaddress emailAddress� �4��
�4 
disp� m  \]�3�3 � �2��
�2 
appr� m  `c�� ��� , O u t l o o k   E x c h a n g e   S e t u p� �1��
�1 
btns� J  fn�� ��� m  fi�� ���  C a n c e l� ��0� m  il�� ���  V e r i f y�0  � �/��.
�/ 
dflt� J  qv�� ��-� m  qt�� ���  V e r i f y�-  �.  � o      �,�, 0 verifyemail verifyEmail� ��� r  ����� n  ����� 1  ���+
�+ 
ttxt� o  ���*�* 0 verifyemail verifyEmail� o      �)�) 0 emailaddress emailAddress� ��(� n ����� I  ���'��&�' 0 writelog writeLog� ��%� b  ����� b  ����� m  ���� ��� > U s e r   v e r i f i e d   e m a i l   a d d r e s s   a s  � o  ���$�$ 0 emailaddress emailAddress� m  ���� ���  .�%  �&  �  f  ���(  �;  �:  � ��� l ���#�"�!�#  �"  �!  � � � Z  �� � = �� o  ���� *0 verifyserveraddress verifyServerAddress m  ���
� boovtrue k  ��  r  ��	 I ���

� .sysodlogaskr        TEXT
 m  �� � f P l e a s e   v e r i f y   y o u r   E x c h a n g e   S e r v e r   n a m e   i s   c o r r e c t . �
� 
dtxt o  ����  0 exchangeserver ExchangeServer �
� 
disp m  ����  �
� 
appr m  �� � , O u t l o o k   E x c h a n g e   S e t u p �
� 
btns J  ��  m  �� �  C a n c e l � m  �� �  V e r i f y�   � �
� 
dflt  J  ��!! "�" m  ��## �$$  V e r i f y�  �  	 o      �� 0 verifyserver verifyServer %&% r  ��'(' n  ��)*) 1  ���
� 
ttxt* o  ���� 0 verifyserver verifyServer( o      ��  0 exchangeserver ExchangeServer& +�+ n ��,-, I  ���.�� 0 writelog writeLog. /�
/ b  ��010 b  ��232 m  ��44 �55 @ U s e r   v e r i f i e d   s e r v e r   a d d r e s s   a s  3 o  ���	�	  0 exchangeserver ExchangeServer1 m  ��66 �77  .�
  �  -  f  ���  �   �    898 l ����  �  �  9 :;: l �<=�  < "  create the Exchange account   = �>> 8   c r e a t e   t h e   E x c h a n g e   a c c o u n t; ?@? l ����  �  �  @ ABA Q  �CDEC k  �FF GHG r  �IJI I ��� K
� .corecrel****      � null�   K ��LM
�� 
koclL m  ��
�� 
EactM ��N��
�� 
prdtN l 	�O����O K  �PP ��QR
�� 
pnamQ b  STS m  UU �VV  M a i l b o x   -  T o  ���� 0 userfullname userFullNameR ��WX
�� 
unmeW b  %YZY o  !���� 0 domainprefix domainPrefixZ o  !$���� 0 usershortname userShortNameX ��[\
�� 
fnam[ o  (+���� 0 userfullname userFullName\ ��]^
�� 
emad] o  .1���� 0 emailaddress emailAddress^ ��_`
�� 
host_ o  49����  0 exchangeserver ExchangeServer` ��ab
�� 
usssa o  <A���� 60 exchangeserverrequiresssl ExchangeServerRequiresSSLb ��cd
�� 
portc o  DI���� .0 exchangeserversslport ExchangeServerSSLPortd ��ef
�� 
ExLSe o  LQ���� "0 directoryserver DirectoryServerf ��gh
�� 
LDAug o  TY���� N0 %directoryserverrequiresauthentication %DirectoryServerRequiresAuthenticationh ��ij
�� 
LDSLi o  \a���� 80 directoryserverrequiresssl DirectoryServerRequiresSSLj ��kl
�� 
ExLPk o  di���� 00 directoryserversslport DirectoryServerSSLPortl ��mn
�� 
LDMXm o  lq���� >0 directoryservermaximumresults DirectoryServerMaximumResultsn ��op
�� 
LDSBo o  ty���� 60 directoryserversearchbase DirectoryServerSearchBasep ��qr
�� 
ExPmq o  |����� *0 downloadheadersonly downloadHeadersOnlyr ��s��
�� 
pBADs o  ������ (0 enableautodiscover enableAutodiscover��  ��  ��  ��  J o      ���� (0 newexchangeaccount newExchangeAccountH t��t n ��uvu I  ����w���� 0 writelog writeLogw x��x m  ��yy �zz H C r e a t e   E x c h a n g e   a c c o u n t :   S u c c e s s f u l .��  ��  v  f  ����  D R      ������
�� .ascrerr ****      � ****��  ��  E k  ��{{ |}| l ����������  ��  ��  } ~~ l ��������  �   something went wrong   � ��� *   s o m e t h i n g   w e n t   w r o n g ��� l ����������  ��  ��  � ��� n ����� I  ��������� 0 writelog writeLog� ���� m  ���� ��� @ C r e a t e   E x c h a n g e   a c c o u n t :   F a i l e d .��  ��  �  f  ��� ��� l ����������  ��  ��  � ��� I ������
�� .sysodlogaskr        TEXT� b  ����� b  ����� b  ����� o  ������ 0 errormessage errorMessage� o  ����
�� 
ret � o  ����
�� 
ret � m  ���� ��� D U n a b l e   t o   c r e a t e   E x c h a n g e   a c c o u n t .� ����
�� 
disp� m  ����
�� stic    � ����
�� 
btns� J  ���� ���� m  ���� ���  O K��  � ����
�� 
dflt� J  ���� ���� m  ���� ���  O K��  � �����
�� 
appr� m  ���� ��� , O u t l o o k   E x c h a n g e   S e t u p��  � ��� R  �������
�� .ascrerr ****      � ****��  � �����
�� 
errn� m  ����������  � ���� l ����������  ��  ��  ��  B ��� l ����������  ��  ��  � ��� l ��������  � e _ The following lines enable Kerberos support if the userKerberos property above is set to true.   � ��� �   T h e   f o l l o w i n g   l i n e s   e n a b l e   K e r b e r o s   s u p p o r t   i f   t h e   u s e r K e r b e r o s   p r o p e r t y   a b o v e   i s   s e t   t o   t r u e .� ��� l ����������  ��  ��  � ��� Z  �n������� = ����� o  ������ 0 usekerberos useKerberos� m  ����
�� boovtrue� Q  �j���� k  ��� ��� r  ���� o  � ���� 0 usekerberos useKerberos� n      ��� 1  ��
�� 
Kerb� o   ���� (0 newexchangeaccount newExchangeAccount� ��� r  	��� o  	���� &0 userkerberosrealm userKerberosRealm� n      ��� 1  ��
�� 
ExGI� o  ���� (0 newexchangeaccount newExchangeAccount� ���� n ��� I  ������� 0 writelog writeLog� ���� m  �� ��� P S e t   K e r b e r o s   a u t h e n t i c a t i o n :   S u c c e s s f u l .��  ��  �  f  ��  � R      ������
�� .ascrerr ****      � ****��  ��  � k  %j�� ��� l %%��������  ��  ��  � ��� l %%������  �   something went wrong   � ��� *   s o m e t h i n g   w e n t   w r o n g� ��� l %%��������  ��  ��  � ��� n %-��� I  &-������� 0 writelog writeLog� ���� m  &)�� ��� H S e t   K e r b e r o s   a u t h e n t i c a t i o n :   F a i l e d .��  ��  �  f  %&� ��� l ..��������  ��  ��  � ��� I .]����
�� .sysodlogaskr        TEXT� b  .;��� b  .7��� b  .5��� o  .3���� 0 errormessage errorMessage� o  34��
�� 
ret � o  56��
�� 
ret � m  7:�� ��� ^ U n a b l e   t o   s e t   E x c h a n g e   a c c o u n t   t o   u s e   K e r b e r o s .� ����
�� 
disp� m  >A��
�� stic    � ����
�� 
btns� J  DI�� ���� m  DG�� ���  O K��  � ��� 
�� 
dflt� J  LQ �� m  LO �  O K��    ���
�� 
appr m  TW � , O u t l o o k   E x c h a n g e   S e t u p�  � 	 R  ^h�~�}

�~ .ascrerr ****      � ****�}  
 �|�{
�| 
errn m  be�z�z���{  	 �y l ii�x�w�v�x  �w  �v  �y  ��  ��  �  l oo�u�t�s�u  �t  �s    Q  ow k  rg  l rr�r�r   M G The Me Contact record is automatically created with the first account.    � �   T h e   M e   C o n t a c t   r e c o r d   i s   a u t o m a t i c a l l y   c r e a t e d   w i t h   t h e   f i r s t   a c c o u n t .  l rr�q�q   a [ Set the first name, last name, email address and other information using Active Directory.    � �   S e t   t h e   f i r s t   n a m e ,   l a s t   n a m e ,   e m a i l   a d d r e s s   a n d   o t h e r   i n f o r m a t i o n   u s i n g   A c t i v e   D i r e c t o r y .   l rr�p�o�n�p  �o  �n    !"! r  r#$# o  ru�m�m 0 userfirstname userFirstName$ n      %&% 1  z~�l
�l 
pFrN& 1  uz�k
�k 
meCn" '(' r  ��)*) o  ���j�j 0 userlastname userLastName* n      +,+ 1  ���i
�i 
pLsN, 1  ���h
�h 
meCn( -.- r  ��/0/ K  ��11 �g23
�g 
radd2 o  ���f�f 0 emailaddress emailAddress3 �e4�d
�e 
type4 m  ���c
�c EATyeWrk�d  0 n      565 1  ���b
�b 
EmAd6 1  ���a
�a 
meCn. 787 r  ��9:9 o  ���`�`  0 userdepartment userDepartment: n      ;<; 1  ���_
�_ 
Dptm< 1  ���^
�^ 
meCn8 =>= r  ��?@? o  ���]�] 0 
useroffice 
userOffice@ n      ABA 1  ���\
�\ 
OficB 1  ���[
�[ 
meCn> CDC r  ��EFE o  ���Z�Z 0 usercompany userCompanyF n      GHG 1  ���Y
�Y 
CmpyH 1  ���X
�X 
meCnD IJI r  ��KLK o  ���W�W 0 userworkphone userWorkPhoneL n      MNM 1  ���V
�V 
bsNmN 1  ���U
�U 
meCnJ OPO r  ��QRQ o  ���T�T 0 
usermobile 
userMobileR n      STS 1  ���S
�S 
mbNmT 1  ���R
�R 
meCnP UVU r  ��WXW o  ���Q�Q 0 userfax userFaxX n      YZY 1  ���P
�P 
bFaxZ 1  ���O
�O 
meCnV [\[ r  �
]^] o  � �N�N 0 	usertitle 	userTitle^ n      _`_ 1  	�M
�M 
pTtl` 1   �L
�L 
meCn\ aba r  cdc o  �K�K 0 
userstreet 
userStreetd n      efe 1  �J
�J 
bStAf 1  �I
�I 
meCnb ghg r  &iji o  �H�H 0 usercity userCityj n      klk 1  !%�G
�G 
bCtyl 1  !�F
�F 
meCnh mnm r  '4opo o  '*�E�E 0 	userstate 	userStatep n      qrq 1  /3�D
�D 
bStar 1  */�C
�C 
meCnn sts r  5Buvu o  58�B�B  0 userpostalcode userPostalCodev n      wxw 1  =A�A
�A 
bZipx 1  8=�@
�@ 
meCnt yzy r  CP{|{ o  CF�?�? 0 usercountry userCountry| n      }~} 1  KO�>
�> 
bCou~ 1  FK�=
�= 
meCnz � r  Q^��� o  QT�<�< 0 userwebpage userWebPage� n      ��� 1  Y]�;
�; 
bsWP� 1  TY�:
�: 
meCn� ��9� n _g��� I  `g�8��7�8 0 writelog writeLog� ��6� m  `c�� ��� X P o p u l a t e   M e   C o n t a c t   i n f o r m a t i o n :   S u c c e s s f u l .�6  �7  �  f  _`�9   R      �5�4�3
�5 .ascrerr ****      � ****�4  �3   n ow��� I  pw�2��1�2 0 writelog writeLog� ��0� m  ps�� ��� P P o p u l a t e   M e   C o n t a c t   i n f o r m a t i o n :   F a i l e d .�0  �1  �  f  op ��� l xx�/�.�-�/  �.  �-  � ��� l xx�,���,  � 0 * Set Outlook to be the default application   � ��� T   S e t   O u t l o o k   t o   b e   t h e   d e f a u l t   a p p l i c a t i o n� ��� l xx�+���+  � ( " for mail, calendars and contacts.   � ��� D   f o r   m a i l ,   c a l e n d a r s   a n d   c o n t a c t s .� ��� l xx�*�)�(�*  �)  �(  � ��� Q  x����� k  {��� ��� r  {���� m  {|�'
�' boovtrue� 1  |��&
�& 
pMSD� ��� r  ����� m  ���%
�% boovtrue� 1  ���$
�$ 
pCSD� ��� r  ����� m  ���#
�# boovtrue� 1  ���"
�" 
pABD� ��!� n ����� I  ��� ���  0 writelog writeLog� ��� m  ���� ��� � S e t   O u t l o o k   a s   d e f a u l t   m a i l ,   c a l e n d a r   a n d   c o n t a c t s   a p p l i c a t i o n :   S u c c e s s f u l .�  �  �  f  ���!  � R      ���
� .ascrerr ****      � ****�  �  � n ����� I  ������ 0 writelog writeLog� ��� m  ���� ��� � S e t   O u t l o o k   a s   d e f a u l t   m a i l ,   c a l e n d a r   a n d   c o n t a c t s   a p p l i c a t i o n :   F a i l e d .�  �  �  f  ��� ��� l ������  �  �  � ��� I �����
� .sysodelanull��� ��� nmbr� m  ���� �  � ��� l ������  �  �  � ��� Q  ������ k  ���� ��� r  ����� m  ���
� boovfals� 1  ���
� 
wkOf� ��� n ����� I  �����
� 0 writelog writeLog� ��	� m  ���� ��� b S e t   M i c r o s o f t   O u t l o o k   t o   w o r k   o n l i n e :   S u c c e s s f u l .�	  �
  �  f  ���  � R      ���
� .ascrerr ****      � ****�  �  � n ����� I  ������ 0 writelog writeLog� ��� m  ���� ��� Z S e t   M i c r o s o f t   O u t l o o k   t o   w o r k   o n l i n e :   F a i l e d .�  �  �  f  ��� ��� l ����� �  �  �   � ��� l ��������  �   We're done.   � ���    W e ' r e   d o n e .� ���� l ����������  ��  ��  ��  [ m  �����                                                                                  OPIM  alis    N  Macintosh HD                   BD ����Microsoft Outlook.app                                          ����            ����  
 cu             Applications  %/:Applications:Microsoft Outlook.app/   ,  M i c r o s o f t   O u t l o o k . a p p    M a c i n t o s h   H D  "Applications/Microsoft Outlook.app  / ��  �t  �s  X ��� l     ��������  ��  ��  � ��� l     ������  � 0 *------------------------------------------   � ��� T - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -� ��� l     ������  �   End account setup   � ��� $   E n d   a c c o u n t   s e t u p� ��� l     ������  � 0 *------------------------------------------   � ��� T - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -� ��� l     ��������  ��  ��  � ��� l     ������  � 0 *------------------------------------------   � ��� T - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -� ��� l     �� ��      Begin script cleanup    � *   B e g i n   s c r i p t   c l e a n u p�  l     ����   0 *------------------------------------------    � T - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 	 l     ��������  ��  ��  	 

 l     ��������  ��  ��    l ������ Q  �� k  ��  I ������
�� .sysoexecTEXT���     TEXT m  �� � � / b i n / r m   $ H O M E / L i b r a r y / L a u n c h A g e n t s / n e t . t a l k i n g m o o s e . O u t l o o k E x c h a n g e S e t u p 5 . p l i s t��   �� I  �������� 0 writelog writeLog �� m  �� � � D e l e t e   O u t l o o k E x c h a n g e S e t u p 5 . p l i s t   f i l e   f r o m   u s e r   L a u n c h A g e n t s   f o l d e r :   S u c c e s s f u l .��  ��  ��   R      ������
�� .ascrerr ****      � ****��  ��   I  �������� 0 writelog writeLog �� m  �� �   � D e l e t e   O u t l o o k E x c h a n g e S e t u p 5 . p l i s t   f i l e   f r o m   u s e r   L a u n c h A g e n t s   f o l d e r :   F a i l e d .��  ��  ��  ��   !"! l     ��������  ��  ��  " #$# l �"%����% Q  �"&'(& k  )) *+* I 	��,��
�� .sysoexecTEXT���     TEXT, m  -- �.. x / b i n / l a u n c h c t l   r e m o v e   n e t . t a l k i n g m o o s e . O u t l o o k E x c h a n g e S e t u p 5��  + /��/ I  
��0���� 0 writelog writeLog0 1��1 m  22 �33 x U n l o a d   O u t l o o k E x c h a n g e S e t u p 5 . p l i s t   l a u n c h   a g e n t :   S u c c e s s f u l .��  ��  ��  ' R      ������
�� .ascrerr ****      � ****��  ��  ( I  "��4���� 0 writelog writeLog4 5��5 m  66 �77 p U n l o a d   O u t l o o k E x c h a n g e S e t u p 5 . p l i s t   l a u n c h   a g e n t :   F a i l e d .��  ��  ��  ��  $ 898 l     ��������  ��  ��  9 :;: l #)<����< I  #)��=���� 0 writelog writeLog= >��> o  $%��
�� 
ret ��  ��  ��  ��  ; ?@? l *0A����A I  *0��B���� 0 writelog writeLogB C��C o  +,��
�� 
ret ��  ��  ��  ��  @ DED l 17F����F I  17��G���� 0 writelog writeLogG H��H o  23��
�� 
ret ��  ��  ��  ��  E IJI l     ��������  ��  ��  J KLK l     ��MN��  M 0 *------------------------------------------   N �OO T - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -L PQP l     ��RS��  R   End script cleanup   S �TT &   E n d   s c r i p t   c l e a n u pQ UVU l     ��WX��  W 0 *------------------------------------------   X �YY T - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -V Z[Z l     ��������  ��  ��  [ \]\ l     ��^_��  ^ 0 *------------------------------------------   _ �`` T - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -] aba l     ��cd��  c   Begin script handlers   d �ee ,   B e g i n   s c r i p t   h a n d l e r sb fgf l     ��hi��  h 0 *------------------------------------------   i �jj T - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -g klk l     ��������  ��  ��  l mnm i   ^ aopo I      ��q���� 0 writelog writeLogq r��r o      ���� 0 
logmessage 
logMessage��  ��  p k     Yss tut r     vwv b     	xyx l    z����z I    ��{|
�� .earsffdralis        afdr{ m     ��
�� afdrcusr| ��}��
�� 
rtyp} m    ��
�� 
TEXT��  ��  ��  y m    ~~ � L L i b r a r y : L o g s : O u t l o o k E x c h a n g e S e t u p 5 . l o gw o      ���� 0 logfile logFileu ��� r    !��� b    ��� b    ��� b    ��� n    ��� 1    ��
�� 
shdt� l   ������ I   ������
�� .misccurdldt    ��� null��  ��  ��  ��  � m    �� ���   � n    ��� 1    ��
�� 
tstr� l   ������ I   ������
�� .misccurdldt    ��� null��  ��  ��  ��  � 1    ��
�� 
tab � o      ���� 0 rightnow rightNow� ��� Z   " 5������ =  " %��� o   " #���� 0 
logmessage 
logMessage� o   # $��
�� 
ret � r   ( +��� o   ( )��
�� 
ret � o      ���� 0 loginfo logInfo��  � r   . 5��� b   . 3��� b   . 1��� o   . /���� 0 rightnow rightNow� o   / 0���� 0 
logmessage 
logMessage� o   1 2��
�� 
ret � o      ���� 0 loginfo logInfo� ��� r   6 B��� I  6 @����
�� .rdwropenshor       file� 4   6 :���
�� 
file� o   8 9���� 0 logfile logFile� ���~
� 
perm� m   ; <�}
�} boovtrue�~  � o      �|�| 0 openlogfile openLogFile� ��� I  C P�{��
�{ .rdwrwritnull���     ****� o   C D�z�z 0 loginfo logInfo� �y��
�y 
refn� o   E F�x�x 0 openlogfile openLogFile� �w��v
�w 
wrat� m   G J�u
�u rdwreof �v  � ��t� I  Q Y�s��r
�s .rdwrclosnull���     ****� 4   Q U�q�
�q 
file� o   S T�p�p 0 logfile logFile�r  �t  n ��� l     �o�n�m�o  �n  �m  � ��� l     �l���l  � 0 *------------------------------------------   � ��� T - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -� ��� l     �k���k  �   End script handlers   � ��� (   E n d   s c r i p t   h a n d l e r s� ��j� l     �i���i  � 0 *------------------------------------------   � ��� T - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -�j       �h��g ?�f�e i�d�c�b�a ��`�_�^�]c�\�[�Z�Y�X�W�V����h  � �U�T�S�R�Q�P�O�N�M�L�K�J�I�H�G�F�E�D�C�B�A�@�?�>�=�<�U 0 usekerberos useKerberos�T  0 exchangeserver ExchangeServer�S 60 exchangeserverrequiresssl ExchangeServerRequiresSSL�R .0 exchangeserversslport ExchangeServerSSLPort�Q "0 directoryserver DirectoryServer�P N0 %directoryserverrequiresauthentication %DirectoryServerRequiresAuthentication�O 80 directoryserverrequiresssl DirectoryServerRequiresSSL�N 00 directoryserversslport DirectoryServerSSLPort�M >0 directoryservermaximumresults DirectoryServerMaximumResults�L 60 directoryserversearchbase DirectoryServerSearchBase�K N0 %getuserinformationfromactivedirectory %getUserInformationFromActiveDirectory�J *0 useemailforusername useEmailForUsername�I 0 
domainname 
domainName�H 0 emailformat emailFormat�G 0 displayname displayName�F 0 domainprefix domainPrefix�E (0 verifyemailaddress verifyEMailAddress�D *0 verifyserveraddress verifyServerAddress�C *0 displaydomainprefix displayDomainPrefix�B *0 downloadheadersonly downloadHeadersOnly�A 20 hideonmycomputerfolders hideOnMyComputerFolders�@ 0 unifiedinbox unifiedInbox�? (0 enableautodiscover enableAutodiscover�> 0 errormessage errorMessage�= 0 writelog writeLog
�< .aevtoappnull  �   � ****
�g boovtrue
�f boovtrue�e�
�d boovtrue
�c boovtrue�b��a�
�` boovtrue
�_ boovfals�^ �] 
�\ boovfals
�[ boovfals
�Z boovfals
�Y boovfals
�X boovfals
�W boovfals
�V boovtrue� �;p�:�9���8�; 0 writelog writeLog�: �7��7 �  �6�6 0 
logmessage 
logMessage�9  � �5�4�3�2�1�5 0 
logmessage 
logMessage�4 0 logfile logFile�3 0 rightnow rightNow�2 0 loginfo logInfo�1 0 openlogfile openLogFile� �0�/�.�-~�,�+��*�)�(�'�&�%�$�#�"�!� �
�0 afdrcusr
�/ 
rtyp
�. 
TEXT
�- .earsffdralis        afdr
�, .misccurdldt    ��� null
�+ 
shdt
�* 
tstr
�) 
tab 
�( 
ret 
�' 
file
�& 
perm
�% .rdwropenshor       file
�$ 
refn
�# 
wrat
�" rdwreof �! 
�  .rdwrwritnull���     ****
� .rdwrclosnull���     ****�8 Z���l �%E�O*j �,�%*j �,%�%E�O��  �E�Y 	��%�%E�O*�/�el E�O���a a  O*�/j � �������
� .aevtoappnull  �   � ****� k    7�� "�� )�� 4�� ]�� d�� m�� v�� �� ��� ��� ��� ��� ��� ��� ��� ��� ��� ��� �� �� �� �� (�� 1�� :�� C�� s�� z�� ��� ��� ��� ��� ��� ��� ��� ��� ��� ��� ��� ��� ��� ��� j�� ��� ��� ��� ��� ��� ��� ��� ��� �   � � �  
   % .		 W

  # : ? D��  �  �  � �� 0 i  �F'�0��bkt}������������&/8Ax����������������������
��	������!(/��E��� M��S��V����\����v�����������������������������-HUi�������	��8EYt�������(5Idq�������%9Ta��u������		��	'	3	I	Q	W	Z	`��������	���	�	�	�	�



O
h
t
�
�
�
�(Zs������&2IQWZ������������#,�����pv���������������������������#��46������U����������������������������������y���������������������������������������������������������������������-26� 0 writelog writeLog
� 
pnam
� 
ret � 0 userfirstname userFirstName� 0 userlastname userLastName�  0 userdepartment userDepartment� 0 
useroffice 
userOffice� 0 usercompany userCompany� 0 userworkphone userWorkPhone� 0 
usermobile 
userMobile� 0 userfax userFax� 0 	usertitle 	userTitle� 0 
userstreet 
userStreet� 0 usercity userCity�
 0 	userstate 	userState�	  0 userpostalcode userPostalCode� 0 usercountry userCountry� 0 userwebpage userWebPage
� .sysoexecTEXT���     TEXT� 0 netbiosdomain netbiosDomain�  �  
� 
disp
� stic    
�  
btns
�� 
dflt
�� 
appr�� 
�� .sysodlogaskr        TEXT
�� 
errn����
�� 
ascr
�� 
txdl�� "0 userinformation userInformation
�� 
cpar
�� .corecnte****       ****
�� 
citm�� 0 emailaddress emailAddress
�� 
cha 
�� 
TEXT�� 0 usershortname userShortName�� 0 userfullname userFullName�� &0 userkerberosrealm userKerberosRealm
�� 
bool
�� .sysosigtsirr   ��� null
�� 
sisn
�� 
siln
�� 
cwor�� 
�� .miscactvnull��� ��� null
�� 
wkOf
�� 
GrpF
�� 
hOMC
�� 
dtxt�� 
�� 0 verifyemail verifyEmail
�� 
ttxt�� 0 verifyserver verifyServer
�� 
kocl
�� 
Eact
�� 
prdt
�� 
unme
�� 
fnam
�� 
emad
�� 
host
�� 
usss
�� 
port
�� 
ExLS
�� 
LDAu
�� 
LDSL
�� 
ExLP
�� 
LDMX
�� 
LDSB
�� 
ExPm
�� 
pBAD�� 
�� .corecrel****      � null�� (0 newexchangeaccount newExchangeAccount
�� 
Kerb
�� 
ExGI
�� 
meCn
�� 
pFrN
�� 
pLsN
�� 
radd
�� 
type
�� EATyeWrk
�� 
EmAd
�� 
Dptm
�� 
Ofic
�� 
Cmpy
�� 
bsNm
�� 
mbNm
�� 
bFax
�� 
pTtl
�� 
bStA
�� 
bCty
�� 
bSta
�� 
bZip
�� 
bCou
�� 
bsWP
�� 
pMSD
�� 
pCSD
�� 
pABD�� 
�� .sysodelanull��� ��� nmbr�8*�k+ O*�)�,%k+ O*�k+ O*�k+ O*�b   %k+ O*�b  %k+ O*�b  %k+ O*�b  %k+ O*�b  %k+ O*�b  %k+ O*�b  %k+ O*�b  %k+ O*�b  %k+ O*�b  	%k+ O*a b  
%k+ O*�k+ Ob  
f  G*a b  %k+ O*a b  %k+ O*a b  %k+ O*a b  %k+ O*�k+ Y hO*a b  %k+ O*a b  %k+ O*a b  %k+ O*a b  %k+ O*a b  %k+ O*a b  %k+ O*a b  %k+ O*a b  %k+ O*�k+ Oa E` Oa E`  Oa !E` "Oa #E` $Oa %E` &Oa 'E` (Oa )E` *Oa +E` ,Oa -E` .Oa /E` 0Oa 1E` 2Oa 3E` 4Oa 5E` 6Oa 7E` 8Oa 9E` :Ob  
e 	5 _a ;j <E` =O*a >_ =%k+ Ob  e  !_ =a ?%Ec  O*a @b  %k+ Y a AEc  O*a Bb  %k+ W JX C Db  �%�%a E%a Fa Ga Ha Ikva Ja Kkva La Ma N OO*a Pk+ O)a Qa RlhO 5a Skv_ Ta U,FOa V_ =%a W%j <E` XO*a Y_ X%k+ OPW JX C Db  �%�%a Z%a Fa Ga Ha [kva Ja \kva La ]a N OO*a ^k+ O)a Qa RlhO{k_ Xa _-j `kh  a akv_ Ta U,FO_ Xa _�/a b L _ Xa _�/a cl/E` dW 2X C Da ekv_ Ta U,FO_ Xa _�k/[a f\[Zl\62a g&E` dY hOa hkv_ Ta U,FO_ Xa _�/a i L _ Xa _�/a cl/E` 8W 2X C Da jkv_ Ta U,FO_ Xa _�k/[a f\[Zl\62a g&E` 8Y hOa kkv_ Ta U,FO_ Xa _�/a l L _ Xa _�/a cl/E` &W 2X C Da mkv_ Ta U,FO_ Xa _�k/[a f\[Zl\62a g&E` &Y hOa nkv_ Ta U,FO_ Xa _�/a o L _ Xa _�/a cl/E` "W 2X C Da pkv_ Ta U,FO_ Xa _�k/[a f\[Zl\62a g&E` "Y hOa qkv_ Ta U,FO_ Xa _�/a r L _ Xa _�/a cl/E` $W 2X C Da skv_ Ta U,FO_ Xa _�k/[a f\[Zl\62a g&E` $Y hOa tkv_ Ta U,FO_ Xa _�/a u L _ Xa _�/a cl/E` vW 2X C Da wkv_ Ta U,FO_ Xa _�k/[a f\[Zl\62a g&E` vY hOa xkv_ Ta U,FO_ Xa _�/a y L _ Xa _�/a cl/E` :W 2X C Da zkv_ Ta U,FO_ Xa _�k/[a f\[Zl\62a g&E` :Y hOa {kv_ Ta U,FO_ Xa _�/a | L _ Xa _�/a cl/E` 2W 2X C Da }kv_ Ta U,FO_ Xa _�k/[a f\[Zl\62a g&E` 2Y hOa ~kv_ Ta U,FO_ Xa _�/a  L _ Xa _�/a cl/E` ,W 2X C Da �kv_ Ta U,FO_ Xa _�k/[a f\[Zl\62a g&E` ,Y hOa �kv_ Ta U,FO_ Xa _�/a � L _ Xa _�/a cl/E` W 2X C Da �kv_ Ta U,FO_ Xa _�k/[a f\[Zl\62a g&E` Y hOa �kv_ Ta U,FO_ Xa _�/a � L _ Xa _�/a cl/E` .W 2X C Da �kv_ Ta U,FO_ Xa _�k/[a f\[Zl\62a g&E` .Y hOa �kv_ Ta U,FO_ Xa _�/a � L _ Xa _�/a cl/E`  W 2X C Da �kv_ Ta U,FO_ Xa _�k/[a f\[Zl\62a g&E`  Y hOa �kv_ Ta U,FO_ Xa _�/a � L _ Xa _�/a cl/E` *W 2X C Da �kv_ Ta U,FO_ Xa _�k/[a f\[Zl\62a g&E` *Y hOa �kv_ Ta U,FO_ Xa _�/a � L _ Xa _�/a cl/E` (W 2X C Da �kv_ Ta U,FO_ Xa _�k/[a f\[Zl\62a g&E` (Y hOa �kv_ Ta U,FO_ Xa _�/a � L _ Xa _�/a cl/E` 6W 2X C Da �kv_ Ta U,FO_ Xa _�k/[a f\[Zl\62a g&E` 6Y hOa �kv_ Ta U,FO_ Xa _�/a � L _ Xa _�/a cl/E` �W 2X C Da �kv_ Ta U,FO_ Xa _�k/[a f\[Zl\62a g&E` �Y hOa �kv_ Ta U,FO_ Xa _�/a � L _ Xa _�/a cl/E` 4W 2X C Da �kv_ Ta U,FO_ Xa _�k/[a f\[Zl\62a g&E` 4Y hOa �kv_ Ta U,FO_ Xa _�/a � L _ Xa _�/a cl/E` 0W 2X C Da �kv_ Ta U,FO_ Xa _�k/[a f\[Zl\62a g&E` 0Y hOP[OY��Oa �a �lv_ Ta U,FO _ Xa cl/E` �W X C DhOa �kv_ Ta U,FO_ da �  Hb  �%�%a �%a Fa Ga Ha �kva Ja �kva La �a N OO*a �k+ O)a Qa RlhY hOPYlb  k 	 b  k a �& s*j �a �,E` vO*j �a �,E` �Oa �_ Ta U,FO_ �a ci/E` O_ �a ck/a �k/E`  Oa �_ Ta U,FO_ a �%_  %a �%b  %E` dOPY�b  k 	 b  l a �& s*j �a �,E` vO*j �a �,E` �Oa �_ Ta U,FO_ �a ck/a �k/E` O_ �a ci/E`  Oa �_ Ta U,FO_ a �%_  %a �%b  %E` dOPYZb  l 	 b  k a �& k*j �a �,E` vO*j �a �,E` �Oa �_ Ta U,FO_ �a ci/E` O_ �a ck/a �k/E`  Oa �_ Ta U,FO_ a �%b  %E` dOPY�b  l 	 b  l a �& k*j �a �,E` vO*j �a �,E` �Oa �_ Ta U,FO_ �a ck/a �k/E` O_ �a ci/E`  Oa �_ Ta U,FO_ a �%b  %E` dOPYXb  m 	 b  k a �& t*j �a �,E` vO*j �a �,E` �Oa �_ Ta U,FO_ �a ci/E` O_ �a ck/a �k/E`  Oa �_ Ta U,FO_ a fk/_  %a �%b  %E` dOPY�b  m 	 b  l a �& t*j �a �,E` vO*j �a �,E` �Oa �_ Ta U,FO_ �a ck/a �k/E` O_ �a ci/E`  Oa �_ Ta U,FO_ a fk/_  %a �%b  %E` dOPYDb  a � 	 b  k a �& k*j �a �,E` vO*j �a �,E` �Oa �_ Ta U,FO_ �a ci/E` O_ �a ck/a �k/E`  Oa �_ Ta U,FO_ va �%b  %E` dOPY �b  a � 	 b  l a �& k*j �a �,E` vO*j �a �,E` �Oa �_ Ta U,FO_ �a ck/a �k/E` O_ �a ci/E`  Oa �_ Ta U,FO_ va �%b  %E` dOPY >b  �%�%a �%a Fa Ga Ha �kva Ja �kva La �a N OO)a Qa RlhOPOb  e  _ dE` vY hO*a �k+ O*a �_ %k+ O*a �_  %k+ O*a �_ d%k+ O*a �_ "%k+ O*a �_ $%k+ O*a �_ &%k+ O*a �_ (%k+ O*a �_ *%k+ O*a �_ ,%k+ O*a �_ .%k+ O*a �_ 0%k+ O*a �_ 2%k+ O*a �_ 4%k+ O*a �_ 6%k+ O*a �_ 8%k+ O*a �_ :%k+ O*�k+ Oa �7*j �O e*a �,FO)a �k+ W X C D)a �k+ O #b  *a �,FO)a �b  %a �%k+ W X C D)a �b  %a �%k+ O #b  *a �,FO)a �b  %a �%k+ W X C D)a �b  %a �%k+ Ob  e  Ra �a �_ da Fka La �a Ha �a �lva Ja �kva � OE` �O_ �a �,E` dO)a �_ d%a �%k+ Y hOb  e  Xa �a �b  a Fka La �a Ha �a �lva Ja �kva � OE` �O_ �a �,Ec  O)a �b  %a �%k+ Y hO �*a �a �a �a_ �%ab  _ v%a_ �a_ dab  ab  ab  ab  a	b  a
b  ab  ab  ab  	ab  ab  aa �E`O)ak+ W LX C D)ak+ Ob  �%�%a%a Fa Ga Hakva Jakva Laa N OO)a Qa RlhOPOb   e  w 'b   _a,FO_ �_a,FO)ak+ W LX C D)ak+ Ob  �%�%a%a Fa Ga Hakva Jakva La a N OO)a Qa RlhOPY hO �_ *a!,a",FO_  *a!,a#,FOa$_ da%a&a �*a!,a',FO_ "*a!,a(,FO_ $*a!,a),FO_ &*a!,a*,FO_ (*a!,a+,FO_ **a!,a,,FO_ ,*a!,a-,FO_ .*a!,a.,FO_ 0*a!,a/,FO_ 2*a!,a0,FO_ 4*a!,a1,FO_ 6*a!,a2,FO_ 8*a!,a3,FO_ :*a!,a4,FO)a5k+ W X C D)a6k+ O %e*a7,FOe*a8,FOe*a9,FO)a:k+ W X C D)a;k+ Oa<j=O f*a �,FO)a>k+ W X C D)a?k+ OPUO a@j <O*aAk+ W X C D*aBk+ O aCj <O*aDk+ W X C D*aEk+ O*�k+ O*�k+ O*�k+ ascr  ��ޭ