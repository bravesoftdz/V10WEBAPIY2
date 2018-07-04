unit UTOTGC;

interface

uses Classes,UTOT,SysUtils,HCtrls,HEnt1,EntGC,HMsgBox,UtilDispGC,
{$IFDEF EAGLCLIENT}
{$ELSE}
    {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}db,
{$ENDIF}
    Controls;
      // mcd 16/09/2002 ajout des fct OnDelete pour libre article ,tiers,contact et établissement  + Fournissuer et piece
Type
    TOT_GCLIBREART1= Class (TOT)
        Function GetTitre :Hstring; override ;
        Procedure OnDeleteRecord ; override ;
        End;
    TOT_GCLIBREART2= Class (TOT)
        Function GetTitre :Hstring; override ;
        Procedure OnDeleteRecord ; override ;
        End;
    TOT_GCLIBREART3= Class (TOT)
        Function GetTitre : Hstring; override ;
        Procedure OnDeleteRecord ; override ;
        End;
    TOT_GCLIBREART4= Class (TOT)
        Function GetTitre : Hstring; override ;
        Procedure OnDeleteRecord ; override ;
        End;
    TOT_GCLIBREART5= Class (TOT)
        Function GetTitre : Hstring; override ;
        Procedure OnDeleteRecord ; override ;
        End;
    TOT_GCLIBREART6= Class (TOT)
        Function GetTitre : Hstring; override ;
        Procedure OnDeleteRecord ; override ;
        End;
    TOT_GCLIBREART7= Class (TOT)
        Function GetTitre : Hstring; override ;
        Procedure OnDeleteRecord ; override ;
        End;
    TOT_GCLIBREART8= Class (TOT)
        Function GetTitre : Hstring; override ;
        Procedure OnDeleteRecord ; override ;
        End;
    TOT_GCLIBREART9= Class (TOT)
        Function GetTitre : Hstring; override ;
        Procedure OnDeleteRecord ; override ;
        End;
    TOT_GCLIBREARTA= Class (TOT)
        Function GetTitre : Hstring; override ;
        Procedure OnDeleteRecord ; override ;
        End;
    TOT_GCLIBRETIERS1= Class (TOT)
        Function GetTitre : Hstring; override ;
        Procedure OnDeleteRecord ; override ;
        End;
    TOT_GCLIBRETIERS2= Class (TOT)
        Function GetTitre : Hstring; override ;
        Procedure OnDeleteRecord ; override ;
        End;
    TOT_GCLIBRETIERS3= Class (TOT)
        Function GetTitre : Hstring; override ;
        Procedure OnDeleteRecord ; override ;
        End;
    TOT_GCLIBRETIERS4= Class (TOT)
        Function GetTitre : Hstring; override ;
        Procedure OnDeleteRecord ; override ;
        End;
    TOT_GCLIBRETIERS5= Class (TOT)
        Function GetTitre : Hstring; override ;
        Procedure OnDeleteRecord ; override ;
        End;
    TOT_GCLIBRETIERS6= Class (TOT)
        Function GetTitre : Hstring; override ;
        Procedure OnDeleteRecord ; override ;
        End;
    TOT_GCLIBRETIERS7= Class (TOT)
        Function GetTitre : Hstring; override ;
        Procedure OnDeleteRecord ; override ;
        End;
    TOT_GCLIBRETIERS8= Class (TOT)
        Function GetTitre : Hstring; override ;
        Procedure OnDeleteRecord ; override ;
        End;
    TOT_GCLIBRETIERS9= Class (TOT)
        Function GetTitre : Hstring; override ;
        Procedure OnDeleteRecord ; override ;
        End;
    TOT_GCLIBRETIERSA= Class (TOT)
        Function GetTitre : Hstring; override ;
        Procedure OnDeleteRecord ; override ;
        End;
    TOT_GCLIBREFOU1= Class (TOT)
        Procedure OnDeleteRecord ; override ;
        End;
    TOT_GCLIBREFOU2= Class (TOT)
        Procedure OnDeleteRecord ; override ;
        End;
    TOT_GCLIBREFOU3= Class (TOT)
        Procedure OnDeleteRecord ; override ;
        End;
    TOT_GCLIBREPIECE1= Class (TOT)
        Procedure OnDeleteRecord ; override ;
        End;
    TOT_GCLIBREPIECE2= Class (TOT)
        Procedure OnDeleteRecord ; override ;
        End;
    TOT_GCLIBREPIECE3= Class (TOT)
        Procedure OnDeleteRecord ; override ;
        End;

    TOT_YYLIBREDEP1= Class (TOT) Function GetTitre : Hstring; override ; End;
    TOT_YYLIBREDEP2= Class (TOT) Function GetTitre : Hstring; override ; End;
    TOT_YYLIBREDEP3= Class (TOT) Function GetTitre :Hstring; override ; End;
    TOT_YYLIBREDEP4= Class (TOT) Function GetTitre :Hstring; override ; End;
    TOT_YYLIBREDEP5= Class (TOT) Function GetTitre :Hstring; override ; End;
    TOT_YYLIBREDEP6= Class (TOT) Function GetTitre :Hstring; override ; End;
    TOT_YYLIBREDEP7= Class (TOT) Function GetTitre :Hstring; override ; End;
    TOT_YYLIBREDEP8= Class (TOT) Function GetTitre :Hstring; override ; End;
    TOT_YYLIBREDEP9= Class (TOT) Function GetTitre :Hstring; override ; End;
    TOT_YYLIBREDEPA= Class (TOT) Function GetTitre :Hstring; override ; End;

{$IFNDEF CCS3}
    TOT_YYLIBREET1= Class (TOT)
        Function GetTitre :Hstring; override ;
        Procedure OnClose ; override ;
        Procedure OnDeleteRecord ; override ;
        End;
    TOT_YYLIBREET2= Class (TOT)
        Function GetTitre :Hstring; override ;
        Procedure OnClose ; override ;
        Procedure OnDeleteRecord ; override ;
        End;
    TOT_YYLIBREET3= Class (TOT)
        Function GetTitre :Hstring; override ;
        Procedure OnClose ; override ;
        Procedure OnDeleteRecord ; override ;
        End;
    TOT_YYLIBREET4= Class (TOT)
        Function GetTitre :Hstring; override ;
        Procedure OnClose ; override ;
        Procedure OnDeleteRecord ; override ;
        End;
    TOT_YYLIBREET5= Class (TOT)
        Function GetTitre :Hstring; override ;
        Procedure OnClose ; override ;
        Procedure OnDeleteRecord ; override ;
        End;
    TOT_YYLIBREET6= Class (TOT)
        Function GetTitre :Hstring; override ;
        Procedure OnClose ; override ;
        Procedure OnDeleteRecord ; override ;
        End;
    TOT_YYLIBREET7= Class (TOT)
        Function GetTitre :Hstring; override ;
        Procedure OnClose ; override ;
        Procedure OnDeleteRecord ; override ;
        End;
    TOT_YYLIBREET8= Class (TOT)
        Function GetTitre :Hstring; override ;
        Procedure OnClose ; override ;
        Procedure OnDeleteRecord ; override ;
        End;
    TOT_YYLIBREET9= Class (TOT)
        Function GetTitre :Hstring; override ;
        Procedure OnClose ; override ;
        Procedure OnDeleteRecord ; override ;
        End;
    TOT_YYLIBREETA= Class (TOT)
        Function GetTitre :Hstring; override ;
        Procedure OnClose ; override ;
        Procedure OnDeleteRecord ; override ;
        End;
{$ENDIF}

    TOT_GCZONECOM= Class (TOT)
        Function GetTitre :Hstring; override ;
        Procedure OnDeleteRecord ; override ;
        End;

    TOT_GCZONELIBRE= Class (TOT)
        Procedure OnClose ; override ;
        Procedure OnAfterUpdateRecord ; override ;
        End;

    TOT_GCLIBFAMILLE= Class (TOT)
        Procedure OnUpdateRecord ; override ;
        End;

    TOT_YYLIBRECON1= Class (TOT)
        Function GetTitre :Hstring; override ;
        Procedure OnDeleteRecord ; override ;
        End;
    TOT_YYLIBRECON2= Class (TOT)
        Function GetTitre :Hstring; override ;
        Procedure OnDeleteRecord ; override ;
        End;
    TOT_YYLIBRECON3= Class (TOT)
        Function GetTitre :Hstring; override ;
        Procedure OnDeleteRecord ; override ;
        End;
    TOT_GCGRILLEDIM1= Class (TOT) Procedure OnDeleteRecord ; override ; End;
    TOT_GCGRILLEDIM2= Class (TOT) Procedure OnDeleteRecord ; override ; End;
    TOT_GCGRILLEDIM3= Class (TOT) Procedure OnDeleteRecord ; override ; End;
    TOT_GCGRILLEDIM4= Class (TOT) Procedure OnDeleteRecord ; override ; End;
    TOT_GCGRILLEDIM5= Class (TOT) Procedure OnDeleteRecord ; override ; End;
    TOT_GCCATEGORIEDIMS3= Class (TOT) Procedure OnUpdateRecord ; override ; End;

    TOT_GCTYPESTATPIECE= Class (TOT)
        Procedure OnUpdateRecord ; override ;
        Procedure OnDeleteRecord ; override ;
        End;

    TOT_GCFAMILLENIV1= Class (TOT)
        Procedure OnDeleteRecord ; override ;
        End;
    TOT_GCFAMILLENIV2= Class (TOT) Procedure OnDeleteRecord ; override ; End;
    TOT_GCFAMILLENIV3= Class (TOT) Procedure OnDeleteRecord ; override ; End;
    TOT_GCFAMILLENIV4= Class (TOT) Procedure OnDeleteRecord ; override ; End;
    TOT_GCFAMILLENIV5= Class (TOT) Procedure OnDeleteRecord ; override ; End;
    TOT_GCFAMILLENIV6= Class (TOT) Procedure OnDeleteRecord ; override ; End;
    TOT_GCFAMILLENIV7= Class (TOT) Procedure OnDeleteRecord ; override ; End;
    TOT_GCFAMILLENIV8= Class (TOT) Procedure OnDeleteRecord ; override ; End;
    TOT_GCCOLLECTION= Class (TOT)  Procedure OnDeleteRecord ; override ; End;
    TOT_GCTARIFARTICLE = Class (TOT) Procedure OnDeleteRecord ; override ; End;
    TOT_TTTARIFCLIENT = Class (TOT) Procedure OnDeleteRecord ; override ; End;
    TOT_GCCOMPTAARTICLE = Class (TOT) Procedure OnDeleteRecord ; override ; End;
    TOT_GCCOMPTATIERS = Class (TOT) Procedure OnDeleteRecord ; override ; End;

{
    TOT_GCTypeMasque = Class (TOT)
        procedure OnUpdateRecord ; override ;
        procedure OnDeleteRecord ; override ;
        procedure OnArgument (stArgument : String ) ; override ;
        END ;
}

{$IFDEF EAGLCLIENT}
// AFAIREEAGL
{$ELSE}
    procedure CreerMasqueAuto(stTypeMasque : string; bMultiEtab : boolean; stMasque : string ; DS : TDataSet) ;
{$ENDIF}


const TexteMessage: array[1..13] of string = (
        {1}   'Cette grille comprend au moins une dimension, elle ne peut pas être supprimée !'
        {2}   ,'Cette table libre comprend au moins une valeur, elle ne peut pas être supprimée !'
        {3}   ,'Cette famille est utilisée dans les articles ou les paramétrages, elle ne peut pas être supprimée !'
        {4}   ,'Cette collection est utilisée dans les articles ou les paramétrages, elle ne peut pas être supprimée !'
        {5}   ,'Cette famille tarifaire est utilisée dans les articles ou les paramétrages, elle ne peut pas être supprimée !'
        {6}   ,'Cette famille tarifaire est utilisée dans les tiers ou les paramétrages, elle ne peut pas être supprimée !'
        {7}   ,'Cette famille comptable tiers  est utilisée dans les tiers ou les paramétrages, elle ne peut pas être supprimée !'
        {8}   ,'Cette famille comptable article est utilisée dans les articles ou les paramétrages, elle ne peut pas être supprimée !'
        {9}   ,'Cette zone commerciale est utilisée dans les commerciaux ou les tiers, elle ne peut pas être supprimée !'
       {10}   ,'Impossible de supprimer cet élément, sa présence est obligatoire !'
       {11}   ,'Des masques utilisent ce type de masques. La suppression de ce type de masques'+#13+#10+
                     'supprimera également les masques de ce type.'+#13+#10+
                     'Confirmez-vous la suppression de ce type de masques ?'
       {12}   ,'Impossible de supprimer cette grille. Elle est utilisée dans les masques'
        {13}  ,'Ce code est utilisé dans une table, il ne peut pas être supprimé !'
);

implementation

Function TOT_GCLIBREART1.GetTitre : Hstring;
BEGIN
	result:=RechDom('GCZONELIBREART','AT1',False)
End;

Function TOT_GCLIBREART2.GetTitre : Hstring;
BEGIN result:=RechDom('GCZONELIBREART','AT2',False) End;

Function TOT_GCLIBREART3.GetTitre : Hstring;
BEGIN result:=RechDom('GCZONELIBREART','AT3',False) End;

Function TOT_GCLIBREART4.GetTitre : Hstring;
BEGIN result:=RechDom('GCZONELIBREART','AT4',False) End;

Function TOT_GCLIBREART5.GetTitre : Hstring;
BEGIN result:=RechDom('GCZONELIBREART','AT5',False) End;

Function TOT_GCLIBREART6.GetTitre : Hstring;
BEGIN result:=RechDom('GCZONELIBREART','AT6',False) End;

Function TOT_GCLIBREART7.GetTitre : Hstring;
BEGIN result:=RechDom('GCZONELIBREART','AT7',False) End;

Function TOT_GCLIBREART8.GetTitre : Hstring;
BEGIN result:=RechDom('GCZONELIBREART','AT8',False) End;

Function TOT_GCLIBREART9.GetTitre : Hstring;
BEGIN result:=RechDom('GCZONELIBREART','AT9',False) End;

Function TOT_GCLIBREARTA.GetTitre : Hstring;
BEGIN result:=RechDom('GCZONELIBREART','ATA',False) End;

Procedure TOT_GCLIBREART1.OnDeleteRecord;
var  SQL : string ;
begin
// Test de l'existence du code dans les tables
SQL:='SELECT GA_LIBREART1 from ARTICLE where GA_LIBREART1="'+getField('YX_CODE')+'"' ;
if ExisteSQL (SQL) then begin LastError:=13; LastErrorMsg:=TexteMessage[LastError]; end ;
SQL:='SELECT GPF_LIBREART1 from PROFILART where GPF_LIBREART1="'+getField('YX_CODE')+'"' ;
if ExisteSQL (SQL) then begin LastError:=13; LastErrorMsg:=TexteMessage[LastError]; end ;
SQL:='SELECT GL_LIBREART1 from LIGNE where GL_LIBREART1="'+getField('YX_CODE')+'"' ;
if ExisteSQL (SQL) then begin LastError:=13; LastErrorMsg:=TexteMessage[LastError]; end ;
SQL:='SELECT BLO_LIBREART1 from LIGNEOUV where BLO_LIBREART1="'+getField('YX_CODE')+'"' ;
if ExisteSQL (SQL) then begin LastError:=13; LastErrorMsg:=TexteMessage[LastError]; end ;
End;

Procedure TOT_GCLIBREART2.OnDeleteRecord;
var  SQL : string ;
begin
// Test de l'existence du code dans les tables
SQL:='SELECT GA_LIBREART2 from ARTICLE where GA_LIBREART2="'+getField('YX_CODE')+'"' ;
if ExisteSQL (SQL) then begin LastError:=13; LastErrorMsg:=TexteMessage[LastError]; end ;
SQL:='SELECT GPF_LIBREART2 from PROFILART where GPF_LIBREART2="'+getField('YX_CODE')+'"' ;
if ExisteSQL (SQL) then begin LastError:=13; LastErrorMsg:=TexteMessage[LastError]; end ;
SQL:='SELECT GL_LIBREART2 from LIGNE where GL_LIBREART2="'+getField('YX_CODE')+'"' ;
if ExisteSQL (SQL) then begin LastError:=13; LastErrorMsg:=TexteMessage[LastError]; end ;
SQL:='SELECT BLO_LIBREART2 from LIGNEOUV where BLO_LIBREART2="'+getField('YX_CODE')+'"' ;
if ExisteSQL (SQL) then begin LastError:=13; LastErrorMsg:=TexteMessage[LastError]; end ;
End;

Procedure TOT_GCLIBREART3.OnDeleteRecord;
var  SQL : string ;
begin
// Test de l'existence du code dans les tables
SQL:='SELECT GA_LIBREART3 from ARTICLE where GA_LIBREART3="'+getField('YX_CODE')+'"' ;
if ExisteSQL (SQL) then begin LastError:=13; LastErrorMsg:=TexteMessage[LastError]; end ;
SQL:='SELECT GPF_LIBREART3 from PROFILART where GPF_LIBREART3="'+getField('YX_CODE')+'"' ;
if ExisteSQL (SQL) then begin LastError:=13; LastErrorMsg:=TexteMessage[LastError]; end ;
SQL:='SELECT GL_LIBREART3 from LIGNE where GL_LIBREART3="'+getField('YX_CODE')+'"' ;
if ExisteSQL (SQL) then begin LastError:=13; LastErrorMsg:=TexteMessage[LastError]; end ;
SQL:='SELECT BLO_LIBREART3 from LIGNEOUV where BLO_LIBREART3="'+getField('YX_CODE')+'"' ;
if ExisteSQL (SQL) then begin LastError:=13; LastErrorMsg:=TexteMessage[LastError]; end ;
End;

Procedure TOT_GCLIBREART4.OnDeleteRecord;
var  SQL : string ;
begin
// Test de l'existence du code dans les tables
SQL:='SELECT GA_LIBREART4 from ARTICLE where GA_LIBREART4="'+getField('YX_CODE')+'"' ;
if ExisteSQL (SQL) then begin LastError:=13; LastErrorMsg:=TexteMessage[LastError]; end ;
SQL:='SELECT GPF_LIBREART4 from PROFILART where GPF_LIBREART4="'+getField('YX_CODE')+'"' ;
if ExisteSQL (SQL) then begin LastError:=13; LastErrorMsg:=TexteMessage[LastError]; end ;
SQL:='SELECT GL_LIBREART4 from LIGNE where GL_LIBREART4="'+getField('YX_CODE')+'"' ;
if ExisteSQL (SQL) then begin LastError:=13; LastErrorMsg:=TexteMessage[LastError]; end ;
SQL:='SELECT BLO_LIBREART4 from LIGNEOUV where BLO_LIBREART4="'+getField('YX_CODE')+'"' ;
if ExisteSQL (SQL) then begin LastError:=13; LastErrorMsg:=TexteMessage[LastError]; end ;
End;

Procedure TOT_GCLIBREART5.OnDeleteRecord;
var  SQL : string ;
begin
// Test de l'existence du code dans les tables
SQL:='SELECT GA_LIBREART5 from ARTICLE where GA_LIBREART5="'+getField('YX_CODE')+'"' ;
if ExisteSQL (SQL) then begin LastError:=13; LastErrorMsg:=TexteMessage[LastError]; end ;
SQL:='SELECT GPF_LIBREART5 from PROFILART where GPF_LIBREART5="'+getField('YX_CODE')+'"' ;
if ExisteSQL (SQL) then begin LastError:=13; LastErrorMsg:=TexteMessage[LastError]; end ;
SQL:='SELECT GL_LIBREART5 from LIGNE where GL_LIBREART5="'+getField('YX_CODE')+'"' ;
if ExisteSQL (SQL) then begin LastError:=13; LastErrorMsg:=TexteMessage[LastError]; end ;
SQL:='SELECT BLO_LIBREART5 from LIGNEOUV where BLO_LIBREART5="'+getField('YX_CODE')+'"' ;
if ExisteSQL (SQL) then begin LastError:=13; LastErrorMsg:=TexteMessage[LastError]; end ;
End;

Procedure TOT_GCLIBREART6.OnDeleteRecord;
var  SQL : string ;
begin
// Test de l'existence du code dans les tables
SQL:='SELECT GA_LIBREART6 from ARTICLE where GA_LIBREART6="'+getField('YX_CODE')+'"' ;
if ExisteSQL (SQL) then begin LastError:=13; LastErrorMsg:=TexteMessage[LastError]; end ;
SQL:='SELECT GPF_LIBREART6 from PROFILART where GPF_LIBREART6="'+getField('YX_CODE')+'"' ;
if ExisteSQL (SQL) then begin LastError:=13; LastErrorMsg:=TexteMessage[LastError]; end ;
SQL:='SELECT GL_LIBREART6 from LIGNE where GL_LIBREART6="'+getField('YX_CODE')+'"' ;
if ExisteSQL (SQL) then begin LastError:=13; LastErrorMsg:=TexteMessage[LastError]; end ;
SQL:='SELECT BLO_LIBREART6 from LIGNEOUV where BLO_LIBREART6="'+getField('YX_CODE')+'"' ;
if ExisteSQL (SQL) then begin LastError:=13; LastErrorMsg:=TexteMessage[LastError]; end ;
End;

Procedure TOT_GCLIBREART7.OnDeleteRecord;
var  SQL : string ;
begin
// Test de l'existence du code dans les tables
SQL:='SELECT GA_LIBREART7 from ARTICLE where GA_LIBREART7="'+getField('YX_CODE')+'"' ;
if ExisteSQL (SQL) then begin LastError:=13; LastErrorMsg:=TexteMessage[LastError]; end ;
SQL:='SELECT GPF_LIBREART7 from PROFILART where GPF_LIBREART7="'+getField('YX_CODE')+'"' ;
if ExisteSQL (SQL) then begin LastError:=13; LastErrorMsg:=TexteMessage[LastError]; end ;
SQL:='SELECT GL_LIBREART7 from LIGNE where GL_LIBREART7="'+getField('YX_CODE')+'"' ;
if ExisteSQL (SQL) then begin LastError:=13; LastErrorMsg:=TexteMessage[LastError]; end ;
SQL:='SELECT BLO_LIBREART7 from LIGNEOUV where BLO_LIBREART7="'+getField('YX_CODE')+'"' ;
if ExisteSQL (SQL) then begin LastError:=13; LastErrorMsg:=TexteMessage[LastError]; end ;
End;

Procedure TOT_GCLIBREART8.OnDeleteRecord;
var  SQL : string ;
begin
// Test de l'existence du code dans les tables
SQL:='SELECT GA_LIBREART8 from ARTICLE where GA_LIBREART8="'+getField('YX_CODE')+'"' ;
if ExisteSQL (SQL) then begin LastError:=13; LastErrorMsg:=TexteMessage[LastError]; end ;
SQL:='SELECT GPF_LIBREART8 from PROFILART where GPF_LIBREART8="'+getField('YX_CODE')+'"' ;
if ExisteSQL (SQL) then begin LastError:=13; LastErrorMsg:=TexteMessage[LastError]; end ;
SQL:='SELECT GL_LIBREART8 from LIGNE where GL_LIBREART8="'+getField('YX_CODE')+'"' ;
if ExisteSQL (SQL) then begin LastError:=13; LastErrorMsg:=TexteMessage[LastError]; end ;
SQL:='SELECT BLO_LIBREART8 from LIGNEOUV where BLO_LIBREART8="'+getField('YX_CODE')+'"' ;
if ExisteSQL (SQL) then begin LastError:=13; LastErrorMsg:=TexteMessage[LastError]; end ;
End;

Procedure TOT_GCLIBREART9.OnDeleteRecord;
var  SQL : string ;
begin
// Test de l'existence du code dans les tables
SQL:='SELECT GA_LIBREART9 from ARTICLE where GA_LIBREART9="'+getField('YX_CODE')+'"' ;
if ExisteSQL (SQL) then begin LastError:=13; LastErrorMsg:=TexteMessage[LastError]; end ;
SQL:='SELECT GPF_LIBREART9 from PROFILART where GPF_LIBREART9="'+getField('YX_CODE')+'"' ;
if ExisteSQL (SQL) then begin LastError:=13; LastErrorMsg:=TexteMessage[LastError]; end ;
SQL:='SELECT GL_LIBREART9 from LIGNE where GL_LIBREART9="'+getField('YX_CODE')+'"' ;
if ExisteSQL (SQL) then begin LastError:=13; LastErrorMsg:=TexteMessage[LastError]; end ;
SQL:='SELECT BLO_LIBREART9 from LIGNEOUV where BLO_LIBREART9="'+getField('YX_CODE')+'"' ;
if ExisteSQL (SQL) then begin LastError:=13; LastErrorMsg:=TexteMessage[LastError]; end ;
End;

Procedure TOT_GCLIBREARTA.OnDeleteRecord;
var  SQL : string ;
begin
// Test de l'existence du code dans les tables
SQL:='SELECT GA_LIBREARTA from ARTICLE where GA_LIBREARTA="'+getField('YX_CODE')+'"' ;
if ExisteSQL (SQL) then begin LastError:=13; LastErrorMsg:=TexteMessage[LastError]; end ;
SQL:='SELECT GPF_LIBREARTA from PROFILART where GPF_LIBREARTA="'+getField('YX_CODE')+'"' ;
if ExisteSQL (SQL) then begin LastError:=13; LastErrorMsg:=TexteMessage[LastError]; end ;
SQL:='SELECT GL_LIBREARTA from LIGNE where GL_LIBREARTA="'+getField('YX_CODE')+'"' ;
if ExisteSQL (SQL) then begin LastError:=13; LastErrorMsg:=TexteMessage[LastError]; end ;
SQL:='SELECT BLO_LIBREARTA from LIGNEOUV where BLO_LIBREARTA="'+getField('YX_CODE')+'"' ;
if ExisteSQL (SQL) then begin LastError:=13; LastErrorMsg:=TexteMessage[LastError]; end ;
End;


Function TOT_GCLIBRETIERS1.GetTitre : Hstring;
BEGIN result:=RechDom('GCZONELIBRETIE','CT1',False) End;

Function TOT_GCLIBRETIERS2.GetTitre : Hstring;
BEGIN result:=RechDom('GCZONELIBRETIE','CT2',False) End;

Function TOT_GCLIBRETIERS3.GetTitre : Hstring;
BEGIN result:=RechDom('GCZONELIBRETIE','CT3',False) End;

Function TOT_GCLIBRETIERS4.GetTitre : Hstring;
BEGIN result:=RechDom('GCZONELIBRETIE','CT4',False) End;

Function TOT_GCLIBRETIERS5.GetTitre : Hstring;
BEGIN result:=RechDom('GCZONELIBRETIE','CT5',False) End;

Function TOT_GCLIBRETIERS6.GetTitre : Hstring;
BEGIN result:=RechDom('GCZONELIBRETIE','CT6',False) End;

Function TOT_GCLIBRETIERS7.GetTitre : Hstring;
BEGIN result:=RechDom('GCZONELIBRETIE','CT7',False) End;

Function TOT_GCLIBRETIERS8.GetTitre : Hstring;
BEGIN result:=RechDom('GCZONELIBRETIE','CT8',False) End;

Function TOT_GCLIBRETIERS9.GetTitre : Hstring;
BEGIN result:=RechDom('GCZONELIBRETIE','CT9',False) End;

Function TOT_GCLIBRETIERSA.GetTitre : Hstring;
BEGIN result:=RechDom('GCZONELIBRETIE','CTA',False) End;

Function TOT_YYLIBREDEP1.GetTitre : Hstring;
BEGIN result:=RechDom('GCZONELIBRE','DT1',False) End;

Function TOT_YYLIBREDEP2.GetTitre : Hstring;
BEGIN result:=RechDom('GCZONELIBRE','DT2',False) End;

Function TOT_YYLIBREDEP3.GetTitre : Hstring;
BEGIN result:=RechDom('GCZONELIBRE','DT3',False) End;

Function TOT_YYLIBREDEP4.GetTitre : Hstring;
BEGIN result:=RechDom('GCZONELIBRE','DT4',False) End;

Function TOT_YYLIBREDEP5.GetTitre : Hstring;
BEGIN result:=RechDom('GCZONELIBRE','DT5',False) End;

Function TOT_YYLIBREDEP6.GetTitre : Hstring;
BEGIN result:=RechDom('GCZONELIBRE','DT6',False) End;

Function TOT_YYLIBREDEP7.GetTitre : Hstring;
BEGIN result:=RechDom('GCZONELIBRE','DT7',False) End;

Function TOT_YYLIBREDEP8.GetTitre : Hstring;
BEGIN result:=RechDom('GCZONELIBRE','DT8',False) End;

Function TOT_YYLIBREDEP9.GetTitre : Hstring;
BEGIN result:=RechDom('GCZONELIBRE','DT9',False) End;

Function TOT_YYLIBREDEPA.GetTitre : Hstring;
BEGIN result:=RechDom('GCZONELIBRE','DTA',False) End;

// Procedure permettant de recopier les données des tables libres des établissements
// dans les tables libres des dépôts
// TypeTabEtab est le type de la tablette (YX_TYPE) de l'établissement

Procedure ChargeTableLibreDepot(TypeTabEtab : string);
var TypeTabDep : string;
begin
// Je créé le type de la tablette du dépôt
TypeTabDep := 'D'+copy(TypeTabEtab,2,2);
// Je supprime tous les enregistrement correspondant au type de la tablette du dépôt
ExecuteSQL('delete from CHOIXEXT where YX_TYPE="'+TypeTabDep+'"');
// j'insere les enregistrements des tables libres de l'établissement dans les tables libres du dépôt
ExecuteSQL('insert into CHOIXEXT (YX_TYPE,YX_CODE,YX_LIBELLE,YX_ABREGE,YX_LIBRE) select "'+TypeTabDep+'" as YX_TYPE,YX_CODE,YX_LIBELLE,YX_ABREGE,YX_LIBRE from CHOIXEXT where YX_TYPE="'+TypeTabEtab+'"');
end;

Procedure TOT_GCLIBRETIERS1.OnDeleteRecord;
var  SQL : string ;
begin
// Test de l'existence du code dans les tables
SQL:='SELECT YTC_TABLELIBRETIERS1 from TIERSCOMPL where YTC_TABLELIBRETIERS1="'+getField('YX_CODE')+'"' ;
if ExisteSQL (SQL) then begin LastError:=13; LastErrorMsg:=TexteMessage[LastError]; end ;
SQL:='SELECT GP_LIBRETIERS1 from PIECE where GP_LIBRETIERS1="'+getField('YX_CODE')+'"' ;
if ExisteSQL (SQL) then begin LastError:=13; LastErrorMsg:=TexteMessage[LastError]; end ;
End;

Procedure TOT_GCLIBRETIERS2.OnDeleteRecord;
var  SQL : string ;
begin
// Test de l'existence du code dans les tables
SQL:='SELECT YTC_TABLELIBRETIERS2 from TIERSCOMPL where YTC_TABLELIBRETIERS2="'+getField('YX_CODE')+'"' ;
if ExisteSQL (SQL) then begin LastError:=13; LastErrorMsg:=TexteMessage[LastError]; end ;
SQL:='SELECT GP_LIBRETIERS2 from PIECE where GP_LIBRETIERS2="'+getField('YX_CODE')+'"' ;
if ExisteSQL (SQL) then begin LastError:=13; LastErrorMsg:=TexteMessage[LastError]; end ;
End;

Procedure TOT_GCLIBRETIERS3.OnDeleteRecord;
var  SQL : string ;
begin
// Test de l'existence du code dans les tables
SQL:='SELECT YTC_TABLELIBRETIERS3 from TIERSCOMPL where YTC_TABLELIBRETIERS3="'+getField('YX_CODE')+'"' ;
if ExisteSQL (SQL) then begin LastError:=13; LastErrorMsg:=TexteMessage[LastError]; end ;
SQL:='SELECT GP_LIBRETIERS3 from PIECE where GP_LIBRETIERS3="'+getField('YX_CODE')+'"' ;
if ExisteSQL (SQL) then begin LastError:=13; LastErrorMsg:=TexteMessage[LastError]; end ;
End;

Procedure TOT_GCLIBRETIERS4.OnDeleteRecord;
var  SQL : string ;
begin
// Test de l'existence du code dans les tables
SQL:='SELECT YTC_TABLELIBRETIERS4 from TIERSCOMPL where YTC_TABLELIBRETIERS4="'+getField('YX_CODE')+'"' ;
if ExisteSQL (SQL) then begin LastError:=13; LastErrorMsg:=TexteMessage[LastError]; end ;
SQL:='SELECT GP_LIBRETIERS4 from PIECE where GP_LIBRETIERS4="'+getField('YX_CODE')+'"' ;
if ExisteSQL (SQL) then begin LastError:=13; LastErrorMsg:=TexteMessage[LastError]; end ;
End;

Procedure TOT_GCLIBRETIERS5.OnDeleteRecord;
var  SQL : string ;
begin
// Test de l'existence du code dans les tables
SQL:='SELECT YTC_TABLELIBRETIERS5 from TIERSCOMPL where YTC_TABLELIBRETIERS5="'+getField('YX_CODE')+'"' ;
if ExisteSQL (SQL) then begin LastError:=13; LastErrorMsg:=TexteMessage[LastError]; end ;
SQL:='SELECT GP_LIBRETIERS5 from PIECE where GP_LIBRETIERS5="'+getField('YX_CODE')+'"' ;
if ExisteSQL (SQL) then begin LastError:=13; LastErrorMsg:=TexteMessage[LastError]; end ;
End;

Procedure TOT_GCLIBRETIERS6.OnDeleteRecord;
var  SQL : string ;
begin
// Test de l'existence du code dans les tables
SQL:='SELECT YTC_TABLELIBRETIERS6 from TIERSCOMPL where YTC_TABLELIBRETIERS6="'+getField('YX_CODE')+'"' ;
if ExisteSQL (SQL) then begin LastError:=13; LastErrorMsg:=TexteMessage[LastError]; end ;
SQL:='SELECT GP_LIBRETIERS6 from PIECE where GP_LIBRETIERS6="'+getField('YX_CODE')+'"' ;
if ExisteSQL (SQL) then begin LastError:=13; LastErrorMsg:=TexteMessage[LastError]; end ;
End;

Procedure TOT_GCLIBRETIERS7.OnDeleteRecord;
var  SQL : string ;
begin
// Test de l'existence du code dans les tables
SQL:='SELECT YTC_TABLELIBRETIERS7 from TIERSCOMPL where YTC_TABLELIBRETIERS7="'+getField('YX_CODE')+'"' ;
if ExisteSQL (SQL) then begin LastError:=13; LastErrorMsg:=TexteMessage[LastError]; end ;
SQL:='SELECT GP_LIBRETIERS7 from PIECE where GP_LIBRETIERS7="'+getField('YX_CODE')+'"' ;
if ExisteSQL (SQL) then begin LastError:=13; LastErrorMsg:=TexteMessage[LastError]; end ;
End;

Procedure TOT_GCLIBRETIERS8.OnDeleteRecord;
var  SQL : string ;
begin
// Test de l'existence du code dans les tables
SQL:='SELECT YTC_TABLELIBRETIERS8 from TIERSCOMPL where YTC_TABLELIBRETIERS8="'+getField('YX_CODE')+'"' ;
if ExisteSQL (SQL) then begin LastError:=13; LastErrorMsg:=TexteMessage[LastError]; end ;
SQL:='SELECT GP_LIBRETIERS8 from PIECE where GP_LIBRETIERS8="'+getField('YX_CODE')+'"' ;
if ExisteSQL (SQL) then begin LastError:=13; LastErrorMsg:=TexteMessage[LastError]; end ;
End;

Procedure TOT_GCLIBRETIERS9.OnDeleteRecord;
var  SQL : string ;
begin
// Test de l'existence du code dans les tables
SQL:='SELECT YTC_TABLELIBRETIERS9 from TIERSCOMPL where YTC_TABLELIBRETIERS9="'+getField('YX_CODE')+'"' ;
if ExisteSQL (SQL) then begin LastError:=13; LastErrorMsg:=TexteMessage[LastError]; end ;
SQL:='SELECT GP_LIBRETIERS9 from PIECE where GP_LIBRETIERS9="'+getField('YX_CODE')+'"' ;
if ExisteSQL (SQL) then begin LastError:=13; LastErrorMsg:=TexteMessage[LastError]; end ;
End;

Procedure TOT_GCLIBRETIERSA.OnDeleteRecord;
var  SQL : string ;
begin
// Test de l'existence du code dans les tables
SQL:='SELECT YTC_TABLELIBRETIERSA from TIERSCOMPL where YTC_TABLELIBRETIERSA="'+getField('YX_CODE')+'"' ;
if ExisteSQL (SQL) then begin LastError:=13; LastErrorMsg:=TexteMessage[LastError]; end ;
SQL:='SELECT GP_LIBRETIERSA from PIECE where GP_LIBRETIERSA="'+getField('YX_CODE')+'"' ;
if ExisteSQL (SQL) then begin LastError:=13; LastErrorMsg:=TexteMessage[LastError]; end ;
End;

Procedure TOT_GCLIBREFOU1.OnDeleteRecord;
var  SQL : string ;
begin
// Test de l'existence du code dans les tables
SQL:='SELECT YTC_TABLELIBREFOU1 from TIERSCOMPL where YTC_TABLELIBREFOU1="'+getField('YX_CODE')+'"' ;
if ExisteSQL (SQL) then begin LastError:=13; LastErrorMsg:=TexteMessage[LastError]; end ;
End;
Procedure TOT_GCLIBREFOU2.OnDeleteRecord;
var  SQL : string ;
begin
// Test de l'existence du code dans les tables
SQL:='SELECT YTC_TABLELIBREFOU2 from TIERSCOMPL where YTC_TABLELIBREFOU2="'+getField('YX_CODE')+'"' ;
if ExisteSQL (SQL) then begin LastError:=13; LastErrorMsg:=TexteMessage[LastError]; end ;
End;
Procedure TOT_GCLIBREFOU3.OnDeleteRecord;
var  SQL : string ;
begin
// Test de l'existence du code dans les tables
SQL:='SELECT YTC_TABLELIBREFOU3 from TIERSCOMPL where YTC_TABLELIBREFOU3="'+getField('YX_CODE')+'"' ;
if ExisteSQL (SQL) then begin LastError:=13; LastErrorMsg:=TexteMessage[LastError]; end ;
End;

Procedure TOT_GCLIBREPIECE1.OnDeleteRecord;
var  SQL : string ;
begin
// Test de l'existence du code dans les tables
SQL:='SELECT GP_LIBREPIECE1 from PIECE where GP_LIBREPIECE1="'+getField('YX_CODE')+'"' ;
if ExisteSQL (SQL) then begin LastError:=13; LastErrorMsg:=TexteMessage[LastError]; end ;
End;
 Procedure TOT_GCLIBREPIECE2.OnDeleteRecord;
var  SQL : string ;
begin
// Test de l'existence du code dans les tables
SQL:='SELECT GP_LIBREPIECE2 from PIECE where GP_LIBREPIECE2="'+getField('YX_CODE')+'"' ;
if ExisteSQL (SQL) then begin LastError:=13; LastErrorMsg:=TexteMessage[LastError]; end ;
End;
Procedure TOT_GCLIBREPIECE3.OnDeleteRecord;
var  SQL : string ;
begin
// Test de l'existence du code dans les tables
SQL:='SELECT GP_LIBREPIECE3 from PIECE where GP_LIBREPIECE3="'+getField('YX_CODE')+'"' ;
if ExisteSQL (SQL) then begin LastError:=13; LastErrorMsg:=TexteMessage[LastError]; end ;
End;



{$IFNDEF CCS3}
Function TOT_YYLIBREET1.GetTitre : Hstring;
BEGIN result:=RechDom('GCZONELIBRE','ET1',False) End;

Procedure TOT_YYLIBREET1.OnClose ;
BEGIN
// Si nous sommes en mono-dépôt, je dois remplir automatiquement les tables libre du dépôt pour qu'elles soient à l'identique des tables libres de l'établissement
if not VH_GC.GCMultiDepots then
   begin
   ChargeTableLibreDepot(GetField('YX_TYPE'));
   end;
{if ExisteSQL('SELECT YX_CODE FROM CHOIXEXT WHERE YX_TYPE="DL1" AND YX_CODE = "'+GetField('YX_CODE')+'"') then
   ExecuteSQL( 'UPDATE CHOIXEXT SET YX_LIBELLE = "'+Getfield ('YX_LIBELLE')+'", YX_ABREGE = "'+GetField('YX_ABREGE')+'" WHERE YX_TYPE = "DL1" and YX_CODE = "'+GetField('YX_CODE')+'"')
else ExecuteSQL( 'INSERT INTO CHOIXEXT VALUES ("DL1","'+GetField('YX_CODE')+'","'+Getfield ('YX_LIBELLE')+'","'+GetField('YX_ABREGE')+'","'+GetField('YX_LIBRE')+'")')}
End;

Function TOT_YYLIBREET2.GetTitre : Hstring;
BEGIN result:=RechDom('GCZONELIBRE','ET2',False) End;

Procedure TOT_YYLIBREET2.OnClose ;
BEGIN
// Si nous sommes en mono-dépôt, je dois remplir automatiquement les tables libre du dépôt pour qu'elles soient à l'identique des tables libres de l'établissement
if not VH_GC.GCMultiDepots then
   begin
   ChargeTableLibreDepot(GetField('YX_TYPE'));
   end;
{if ExisteSQL('SELECT YX_CODE FROM CHOIXEXT WHERE YX_TYPE="DL2" AND YX_CODE = "'+GetField('YX_CODE')+'"') then
   ExecuteSQL( 'UPDATE CHOIXEXT SET YX_LIBELLE = "'+Getfield ('YX_LIBELLE')+'", YX_ABREGE = "'+GetField('YX_ABREGE')+'" WHERE YX_TYPE = "DL2" and YX_CODE = "'+GetField('YX_CODE')+'"')
else ExecuteSQL( 'INSERT INTO CHOIXEXT VALUES ("DL2","'+GetField('YX_CODE')+'","'+Getfield ('YX_LIBELLE')+'","'+GetField('YX_ABREGE')+'","'+GetField('YX_LIBRE')+'")')}
End;

Function TOT_YYLIBREET3.GetTitre : Hstring;
BEGIN result:=RechDom('GCZONELIBRE','ET3',False) End;

Procedure TOT_YYLIBREET3.OnClose ;
BEGIN
// Si nous sommes en mono-dépôt, je dois remplir automatiquement les tables libre du dépôt pour qu'elles soient à l'identique des tables libres de l'établissement
if not VH_GC.GCMultiDepots then
   begin
   ChargeTableLibreDepot(GetField('YX_TYPE'));
   end;
{if ExisteSQL('SELECT YX_CODE FROM CHOIXEXT WHERE YX_TYPE="DL3" AND YX_CODE = "'+GetField('YX_CODE')+'"') then
   ExecuteSQL( 'UPDATE CHOIXEXT SET YX_LIBELLE = "'+Getfield ('YX_LIBELLE')+'", YX_ABREGE = "'+GetField('YX_ABREGE')+'" WHERE YX_TYPE = "DL3" and YX_CODE = "'+GetField('YX_CODE')+'"')
else ExecuteSQL( 'INSERT INTO CHOIXEXT VALUES ("DL3","'+GetField('YX_CODE')+'","'+Getfield ('YX_LIBELLE')+'","'+GetField('YX_ABREGE')+'","'+GetField('YX_LIBRE')+'")')}
End;

Function TOT_YYLIBREET4.GetTitre : Hstring;
BEGIN result:=RechDom('GCZONELIBRE','ET4',False) End;

Procedure TOT_YYLIBREET4.OnClose ;
BEGIN
// Si nous sommes en mono-dépôt, je dois remplir automatiquement les tables libre du dépôt pour qu'elles soient à l'identique des tables libres de l'établissement
if not VH_GC.GCMultiDepots then
   begin
   ChargeTableLibreDepot(GetField('YX_TYPE'));
   end;
{if ExisteSQL('SELECT YX_CODE FROM CHOIXEXT WHERE YX_TYPE="DL4" AND YX_CODE = "'+GetField('YX_CODE')+'"') then
   ExecuteSQL( 'UPDATE CHOIXEXT SET YX_LIBELLE = "'+Getfield ('YX_LIBELLE')+'", YX_ABREGE = "'+GetField('YX_ABREGE')+'" WHERE YX_TYPE = "DL4" and YX_CODE = "'+GetField('YX_CODE')+'"')
else ExecuteSQL( 'INSERT INTO CHOIXEXT VALUES ("DL4","'+GetField('YX_CODE')+'","'+Getfield ('YX_LIBELLE')+'","'+GetField('YX_ABREGE')+'","'+GetField('YX_LIBRE')+'")')}
End;

Function TOT_YYLIBREET5.GetTitre : Hstring;
BEGIN result:=RechDom('GCZONELIBRE','ET5',False) End;

Procedure TOT_YYLIBREET5.OnClose ;
BEGIN
// Si nous sommes en mono-dépôt, je dois remplir automatiquement les tables libre du dépôt pour qu'elles soient à l'identique des tables libres de l'établissement
if not VH_GC.GCMultiDepots then
   begin
   ChargeTableLibreDepot(GetField('YX_TYPE'));
   end;
{if ExisteSQL('SELECT YX_CODE FROM CHOIXEXT WHERE YX_TYPE="DL5" AND YX_CODE = "'+GetField('YX_CODE')+'"') then
   ExecuteSQL( 'UPDATE CHOIXEXT SET YX_LIBELLE = "'+Getfield ('YX_LIBELLE')+'", YX_ABREGE = "'+GetField('YX_ABREGE')+'" WHERE YX_TYPE = "DL5" and YX_CODE = "'+GetField('YX_CODE')+'"')
else ExecuteSQL( 'INSERT INTO CHOIXEXT VALUES ("DL5","'+GetField('YX_CODE')+'","'+Getfield ('YX_LIBELLE')+'","'+GetField('YX_ABREGE')+'","'+GetField('YX_LIBRE')+'")')}
End;

Function TOT_YYLIBREET6.GetTitre : Hstring;
BEGIN result:=RechDom('GCZONELIBRE','ET6',False) End;

Procedure TOT_YYLIBREET6.OnClose ;
BEGIN
// Si nous sommes en mono-dépôt, je dois remplir automatiquement les tables libre du dépôt pour qu'elles soient à l'identique des tables libres de l'établissement
if not VH_GC.GCMultiDepots then
   begin
   ChargeTableLibreDepot(GetField('YX_TYPE'));
   end;
{if ExisteSQL('SELECT YX_CODE FROM CHOIXEXT WHERE YX_TYPE="DL6" AND YX_CODE = "'+GetField('YX_CODE')+'"') then
   ExecuteSQL( 'UPDATE CHOIXEXT SET YX_LIBELLE = "'+Getfield ('YX_LIBELLE')+'", YX_ABREGE = "'+GetField('YX_ABREGE')+'" WHERE YX_TYPE = "DL6" and YX_CODE = "'+GetField('YX_CODE')+'"')
else ExecuteSQL( 'INSERT INTO CHOIXEXT VALUES ("DL6","'+GetField('YX_CODE')+'","'+Getfield ('YX_LIBELLE')+'","'+GetField('YX_ABREGE')+'","'+GetField('YX_LIBRE')+'")')}
End;

Function TOT_YYLIBREET7.GetTitre : Hstring;
BEGIN result:=RechDom('GCZONELIBRE','ET7',False) End;

Procedure TOT_YYLIBREET7.OnClose ;
BEGIN
// Si nous sommes en mono-dépôt, je dois remplir automatiquement les tables libre du dépôt pour qu'elles soient à l'identique des tables libres de l'établissement
if not VH_GC.GCMultiDepots then
   begin
   ChargeTableLibreDepot(GetField('YX_TYPE'));
   end;
{if ExisteSQL('SELECT YX_CODE FROM CHOIXEXT WHERE YX_TYPE="DL7" AND YX_CODE = "'+GetField('YX_CODE')+'"') then
   ExecuteSQL( 'UPDATE CHOIXEXT SET YX_LIBELLE = "'+Getfield ('YX_LIBELLE')+'", YX_ABREGE = "'+GetField('YX_ABREGE')+'" WHERE YX_TYPE = "DL7" and YX_CODE = "'+GetField('YX_CODE')+'"')
else ExecuteSQL( 'INSERT INTO CHOIXEXT VALUES ("DL7","'+GetField('YX_CODE')+'","'+Getfield ('YX_LIBELLE')+'","'+GetField('YX_ABREGE')+'","'+GetField('YX_LIBRE')+'")')}
End;

Function TOT_YYLIBREET8.GetTitre : Hstring;
BEGIN result:=RechDom('GCZONELIBRE','ET8',False) End;

Procedure TOT_YYLIBREET8.OnClose ;
BEGIN
// Si nous sommes en mono-dépôt, je dois remplir automatiquement les tables libre du dépôt pour qu'elles soient à l'identique des tables libres de l'établissement
if not VH_GC.GCMultiDepots then
   begin
   ChargeTableLibreDepot(GetField('YX_TYPE'));
   end;
{if ExisteSQL('SELECT YX_CODE FROM CHOIXEXT WHERE YX_TYPE="DL8" AND YX_CODE = "'+GetField('YX_CODE')+'"') then
   ExecuteSQL( 'UPDATE CHOIXEXT SET YX_LIBELLE = "'+Getfield ('YX_LIBELLE')+'", YX_ABREGE = "'+GetField('YX_ABREGE')+'" WHERE YX_TYPE = "DL8" and YX_CODE = "'+GetField('YX_CODE')+'"')
else ExecuteSQL( 'INSERT INTO CHOIXEXT VALUES ("DL8","'+GetField('YX_CODE')+'","'+Getfield ('YX_LIBELLE')+'","'+GetField('YX_ABREGE')+'","'+GetField('YX_LIBRE')+'")')}
End;

Function TOT_YYLIBREET9.GetTitre : Hstring;
BEGIN result:=RechDom('GCZONELIBRE','ET9',False) End;

Procedure TOT_YYLIBREET9.OnClose ;
BEGIN
// Si nous sommes en mono-dépôt, je dois remplir automatiquement les tables libre du dépôt pour qu'elles soient à l'identique des tables libres de l'établissement
if not VH_GC.GCMultiDepots then
   begin
   ChargeTableLibreDepot(GetField('YX_TYPE'));
   end;
{if ExisteSQL('SELECT YX_CODE FROM CHOIXEXT WHERE YX_TYPE="DL9" AND YX_CODE = "'+GetField('YX_CODE')+'"') then
   ExecuteSQL( 'UPDATE CHOIXEXT SET YX_LIBELLE = "'+Getfield ('YX_LIBELLE')+'", YX_ABREGE = "'+GetField('YX_ABREGE')+'" WHERE YX_TYPE = "DL9" and YX_CODE = "'+GetField('YX_CODE')+'"')
else ExecuteSQL( 'INSERT INTO CHOIXEXT VALUES ("DL9","'+GetField('YX_CODE')+'","'+Getfield ('YX_LIBELLE')+'","'+GetField('YX_ABREGE')+'","'+GetField('YX_LIBRE')+'")')}
End;

Function TOT_YYLIBREETA.GetTitre : Hstring;
BEGIN result:=RechDom('GCZONELIBRE','ETA',False) End;

Procedure TOT_YYLIBREETA.OnClose ;
BEGIN
// Si nous sommes en mono-dépôt, je dois remplir automatiquement les tables libre du dépôt pour qu'elles soient à l'identique des tables libres de l'établissement
if not VH_GC.GCMultiDepots then
   begin
   ChargeTableLibreDepot(GetField('YX_TYPE'));
   end;
{if ExisteSQL('SELECT YX_CODE FROM CHOIXEXT WHERE YX_TYPE="DLA" AND YX_CODE = "'+GetField('YX_CODE')+'"') then
   ExecuteSQL( 'UPDATE CHOIXEXT SET YX_LIBELLE = "'+Getfield ('YX_LIBELLE')+'", YX_ABREGE = "'+GetField('YX_ABREGE')+'" WHERE YX_TYPE = "DLA" and YX_CODE = "'+GetField('YX_CODE')+'"')
else ExecuteSQL( 'INSERT INTO CHOIXEXT VALUES ("DLA","'+GetField('YX_CODE')+'","'+Getfield ('YX_LIBELLE')+'","'+GetField('YX_ABREGE')+'","'+GetField('YX_LIBRE')+'")')}
End;

Procedure TOT_YYLIBREET1.OnDeleteRecord;
var  SQL : string ;
begin
// Test de l'existence du code dans la table
SQL:='SELECT ET_LIBREET1 from ETABLISS where ET_LIBREET1="'+getField('YX_CODE')+'"' ;
if ExisteSQL (SQL) then begin LastError:=13; LastErrorMsg:=TexteMessage[LastError]; end ;
End;

Procedure TOT_YYLIBREET2.OnDeleteRecord;
var  SQL : string ;
begin
// Test de l'existence du code dans la table
SQL:='SELECT ET_LIBREET2 from ETABLISS where ET_LIBREET2="'+getField('YX_CODE')+'"' ;
if ExisteSQL (SQL) then begin LastError:=13; LastErrorMsg:=TexteMessage[LastError]; end ;
End;
Procedure TOT_YYLIBREET3.OnDeleteRecord;
var  SQL : string ;
begin
// Test de l'existence du code dans la table
SQL:='SELECT ET_LIBREET3 from ETABLISS where ET_LIBREET3="'+getField('YX_CODE')+'"' ;
if ExisteSQL (SQL) then begin LastError:=13; LastErrorMsg:=TexteMessage[LastError]; end ;
End;
Procedure TOT_YYLIBREET4.OnDeleteRecord;
var  SQL : string ;
begin
// Test de l'existence du code dans la table
SQL:='SELECT ET_LIBREET4 from ETABLISS where ET_LIBREET4="'+getField('YX_CODE')+'"' ;
if ExisteSQL (SQL) then begin LastError:=13; LastErrorMsg:=TexteMessage[LastError]; end ;
End;
Procedure TOT_YYLIBREET5.OnDeleteRecord;
var  SQL : string ;
begin
// Test de l'existence du code dans la table
SQL:='SELECT ET_LIBREET5 from ETABLISS where ET_LIBREET5="'+getField('YX_CODE')+'"' ;
if ExisteSQL (SQL) then begin LastError:=13; LastErrorMsg:=TexteMessage[LastError]; end ;
End;
Procedure TOT_YYLIBREET6.OnDeleteRecord;
var  SQL : string ;
begin
// Test de l'existence du code dans la table
SQL:='SELECT ET_LIBREET6 from ETABLISS where ET_LIBREET6="'+getField('YX_CODE')+'"' ;
if ExisteSQL (SQL) then begin LastError:=13; LastErrorMsg:=TexteMessage[LastError]; end ;
End;
Procedure TOT_YYLIBREET7.OnDeleteRecord;
var  SQL : string ;
begin
// Test de l'existence du code dans la table
SQL:='SELECT ET_LIBREET7 from ETABLISS where ET_LIBREET7="'+getField('YX_CODE')+'"' ;
if ExisteSQL (SQL) then begin LastError:=13; LastErrorMsg:=TexteMessage[LastError]; end ;
End;
Procedure TOT_YYLIBREET8.OnDeleteRecord;
var  SQL : string ;
begin
// Test de l'existence du code dans la table
SQL:='SELECT ET_LIBREET8 from ETABLISS where ET_LIBREET8="'+getField('YX_CODE')+'"' ;
if ExisteSQL (SQL) then begin LastError:=13; LastErrorMsg:=TexteMessage[LastError]; end ;
End;
Procedure TOT_YYLIBREET9.OnDeleteRecord;
var  SQL : string ;
begin
// Test de l'existence du code dans la table
SQL:='SELECT ET_LIBREET9 from ETABLISS where ET_LIBREET9="'+getField('YX_CODE')+'"' ;
if ExisteSQL (SQL) then begin LastError:=13; LastErrorMsg:=TexteMessage[LastError]; end ;
End;
Procedure TOT_YYLIBREETA.OnDeleteRecord;
var  SQL : string ;
begin
// Test de l'existence du code dans la table
SQL:='SELECT ET_LIBREETA from ETABLISS where ET_LIBREETA="'+getField('YX_CODE')+'"' ;
if ExisteSQL (SQL) then begin LastError:=13; LastErrorMsg:=TexteMessage[LastError]; end ;
End;
{$ENDIF}

Function TOT_YYLIBRECON1.GetTitre : Hstring;
BEGIN result:=RechDom('GCZONELIBRE','BT1',False) End;

Function TOT_YYLIBRECON2.GetTitre : Hstring;
BEGIN result:=RechDom('GCZONELIBRE','BT2',False) End;

Function TOT_YYLIBRECON3.GetTitre : Hstring;
BEGIN result:=RechDom('GCZONELIBRE','BT3',False) End;

Procedure TOT_YYLIBRECON1.OnDeleteRecord;
var  SQL : string ;
begin
// Test de l'existence du code dans la table
SQL:='SELECT C_LIBRECONTACT1 from CONTACT where C_LIBRECONTACT1="'+getField('YX_CODE')+'"' ;
if ExisteSQL (SQL) then begin LastError:=13; LastErrorMsg:=TexteMessage[LastError]; end ;
End;

Procedure TOT_YYLIBRECON2.OnDeleteRecord;
var  SQL : string ;
begin
// Test de l'existence du code dans la table
SQL:='SELECT C_LIBRECONTACT2 from CONTACT where C_LIBRECONTACT2="'+getField('YX_CODE')+'"' ;
if ExisteSQL (SQL) then begin LastError:=13; LastErrorMsg:=TexteMessage[LastError]; end ;
End;

Procedure TOT_YYLIBRECON3.OnDeleteRecord;
var  SQL : string ;
begin
// Test de l'existence du code dans la table
SQL:='SELECT C_LIBRECONTACT3 from CONTACT where C_LIBRECONTACT3="'+getField('YX_CODE')+'"' ;
if ExisteSQL (SQL) then begin LastError:=13; LastErrorMsg:=TexteMessage[LastError]; end ;
End;

Function TOT_GCZONECOM.GetTitre : Hstring;
BEGIN
If ctxScot in V_PGI.PGIContexte  then result:='Secteur de l''apporteur'
                                 else result:=TTToLibelle('GCZONECOM' ) ;
End;

Function ExisteZoneCom ( Code : String ) : Boolean ;
BEGIN
Result:=True ;
if ExisteSQL('SELECT T_ZONECOM FROM TIERS WHERE T_ZONECOM="'+Code+'"') then Exit ;
if ExisteSQL('SELECT GCL_ZONECOM FROM COMMERCIAL WHERE GCL_ZONECOM="'+Code+'"') then Exit ;
Result:=False ;
END ;

Procedure TOT_GCZONECOM.OnDeleteRecord ;
var  Code : string ;
begin
Code:=GetField('CC_CODE') ;
if ExisteZoneCom(Code) then begin LastError:=9; LastErrorMsg:=TexteMessage[LastError]; end ;
END ;

// Procedure permettant de recopier les données des tables libres des établissements
// dans les tables libres des dépôts
// TypeTabEtab est le type de la tablette (YX_TYPE) de l'établissement

Procedure ChargeTitreZoneLibreDepot(CodeTabEtab : string);
var CodeTabDep : string;
begin
// Je créé le type de la tablette du dépôt
CodeTabDep := 'D'+copy(CodeTabEtab,2,1);
// Je supprime tous les enregistrement correspondant au type de la tablette du dépôt
ExecuteSQL('delete from CHOIXCOD where CC_TYPE="ZLI" and CC_CODE like"'+CodeTabDep+'%"');
// j'insere les enregistrements des tables libres de l'établissement dans les tables libres du dépôt
ExecuteSQL('insert into CHOIXCOD (CC_TYPE,CC_CODE,CC_LIBELLE,CC_ABREGE,CC_LIBRE) select "ZLI" as CC_TYPE,"D"||right(CC_CODE,2) as CC_CODE,CC_LIBELLE,CC_ABREGE,CC_LIBRE from CHOIXCOD where CC_TYPE="ZLI" and CC_CODE like "'+copy(CodeTabEtab,1,2)+'%"');
end;

Procedure TOT_GCZONELIBRE.OnClose ;
Begin
if not VH_GC.GCMultiDepots then
   begin
   ChargeTitreZoneLibreDepot(GetField('CC_CODE'));
   //ExecuteSQL( 'UPDATE CHOIXCOD SET CC_LIBELLE = "'+Getfield ('CC_LIBELLE')+'", CC_ABREGE = "'+GetField('CC_ABREGE')+'" WHERE CC_TYPE = "ZLI" and CC_CODE = "'+Code+'"') ;
   end;
End;

Procedure TOT_GCZONELIBRE.OnAfterUpdateRecord;
Begin
AvertirTable ('GCZONELIBRE');
AvertirTable ('GCZONELIBREART');
AvertirTable ('GCZONELIBRETIE');
TraiteMenusMode(105);
End;

Procedure TOT_GCLIBFAMILLE.OnUpdateRecord;
Begin
// Forcer le libellé personalisé des familles niv 1, niv 2 et niv 3 dans les différentes tablettes
// Tablette Regroupement édition Ctrl Marge
ExecuteSQL( 'UPDATE COMMUN SET CO_LIBELLE = "'+Getfield ('CC_LIBELLE')+'" WHERE CO_TYPE = "GCM" and CO_CODE = "'+Getfield ('CC_CODE')+'"') ;
AvertirTable ('GCGROUPCTRLMARGE');
// Tablette Regroupement édition consommation
ExecuteSQL( 'UPDATE COMMUN SET CO_LIBELLE = "'+Getfield ('CC_LIBELLE')+'" WHERE CO_TYPE = "GGC" and CO_CODE = "'+Getfield ('CC_CODE')+'"') ;
AvertirTable ('GCGROUPCONSO');
// Tablette Regroupement édition inventaire
ExecuteSQL( 'UPDATE COMMUN SET CO_LIBELLE = "'+Getfield ('CC_LIBELLE')+'" WHERE CO_TYPE = "GGI" and CO_CODE = "'+Getfield ('CC_CODE')+'"') ;
AvertirTable ('GCGROUPINV');
// Tablette Regroupement édition inventaire permanent
ExecuteSQL( 'UPDATE COMMUN SET CO_LIBELLE = "'+Getfield ('CC_LIBELLE')+'" WHERE CO_TYPE = "GGV" and CO_CODE = "'+Getfield ('CC_CODE')+'"') ;
AvertirTable ('GCGROUPINVPERM');
// Tablette Regroupement édition des Palmarès
ExecuteSQL( 'UPDATE COMMUN SET CO_LIBELLE = "'+Getfield ('CC_LIBELLE')+'" WHERE CO_TYPE = "GPL" and CO_CODE = "'+Getfield ('CC_CODE')+'"') ;
AvertirTable ('GCGROUPPALM');
// Tablette Regroupement édition inventaire permanent
ExecuteSQL( 'UPDATE COMMUN SET CO_LIBELLE = "'+Getfield ('CC_LIBELLE')+'" WHERE CO_TYPE = "GST" and CO_CODE = "'+Getfield ('CC_CODE')+'"') ;
AvertirTable ('GCGROUPSTA');
End;

Procedure TOT_GCGRILLEDIM1.OnDeleteRecord;
var  SQL : string ;
begin
// Test de l'existence d'une dimension dans la grille
SQL:='SELECT GDI_CODEDIM from DIMENSION where GDI_TYPEDIM="DI1" and GDI_GRILLEDIM="'+getField('CC_CODE')+'"' ;
if ExisteSQL (SQL) then begin LastError:=1; LastErrorMsg:=TexteMessage[LastError]; end ;
// Test de l'existence de la grille dans un masque
SQL:='SELECT GDM_MASQUE FROM DIMMASQUE where GDM_TYPE1="'+getField('CC_CODE')+'"' ;
if ExisteSQL (SQL) then begin LastError:=12; LastErrorMsg:=TexteMessage[LastError]; end ;
End;

Procedure TOT_GCGRILLEDIM2.OnDeleteRecord;
var  SQL : string ;
begin
// Test de l'existence d'une dimension dans la grille
SQL:='SELECT GDI_CODEDIM from DIMENSION where GDI_TYPEDIM="DI2" and GDI_GRILLEDIM="'+getField('CC_CODE')+'"' ;
if ExisteSQL (SQL) then begin LastError:=1; LastErrorMsg:=TexteMessage[LastError]; end ;
// Test de l'existence de la grille dans un masque
SQL:='SELECT GDM_MASQUE FROM DIMMASQUE where GDM_TYPE2="'+getField('CC_CODE')+'"' ;
if ExisteSQL (SQL) then begin LastError:=12; LastErrorMsg:=TexteMessage[LastError]; end ;
End;

Procedure TOT_GCGRILLEDIM3.OnDeleteRecord;
var  SQL : string ;
begin
// Test de l'existence d'une dimension dans la grille
SQL:='SELECT GDI_CODEDIM from DIMENSION where GDI_TYPEDIM="DI3" and GDI_GRILLEDIM="'+getField('CC_CODE')+'"' ;
if ExisteSQL (SQL) then begin LastError:=1; LastErrorMsg:=TexteMessage[LastError]; end ;
// Test de l'existence de la grille dans un masque
SQL:='SELECT GDM_MASQUE FROM DIMMASQUE where GDM_TYPE3="'+getField('CC_CODE')+'"' ;
if ExisteSQL (SQL) then begin LastError:=12; LastErrorMsg:=TexteMessage[LastError]; end ;
End;

Procedure TOT_GCGRILLEDIM4.OnDeleteRecord;
var  SQL : string ;
begin
// Test de l'existence d'une dimension dans la grille
SQL:='SELECT GDI_CODEDIM from DIMENSION where GDI_TYPEDIM="DI4" and GDI_GRILLEDIM="'+getField('CC_CODE')+'"' ;
if ExisteSQL (SQL) then begin LastError:=1; LastErrorMsg:=TexteMessage[LastError]; end ;
// Test de l'existence de la grille dans un masque
SQL:='SELECT GDM_MASQUE FROM DIMMASQUE where GDM_TYPE4="'+getField('CC_CODE')+'"' ;
if ExisteSQL (SQL) then begin LastError:=12; LastErrorMsg:=TexteMessage[LastError]; end ;
End;

Procedure TOT_GCGRILLEDIM5.OnDeleteRecord;
var  SQL : string ;
begin
// Test de l'existence d'une dimension dans la grille
SQL:='SELECT GDI_CODEDIM from DIMENSION where GDI_TYPEDIM="DI5" and GDI_GRILLEDIM="'+getField('CC_CODE')+'"' ;
if ExisteSQL (SQL) then begin LastError:=1; LastErrorMsg:=TexteMessage[LastError]; end ;
// Test de l'existence de la grille dans un masque
SQL:='SELECT GDM_MASQUE FROM DIMMASQUE where GDM_TYPE5="'+getField('CC_CODE')+'"' ;
if ExisteSQL (SQL) then begin LastError:=12; LastErrorMsg:=TexteMessage[LastError]; end ;
End;

Procedure TOT_GCCATEGORIEDIMS3.OnUpdateRecord;
begin
// Forcer le libellé ".-" sur les dimensions niv 4 et niv 5, qui ne doivent pas être utilisé
// sur la Gamme MODE S3
ExecuteSQL( 'UPDATE CHOIXCOD SET CC_LIBELLE = ".-" WHERE CC_TYPE = "DIM" and CC_CODE > "DI3"') ;
AvertirTable ('GCCATEGORIEDIM');
End;

Procedure TOT_GCTYPESTATPIECE.OnUpdateRecord;
Var TailleCode : integer;
    NewCode : String;
begin
TailleCode:=length(getField('CC_CODE'));
if (TailleCode>0) and (TailleCode<3) then
  begin
  NewCode := getField('CC_CODE') + StringOfChar('.',3-TailleCode);
  SetField('CC_CODE',NewCode);
  end;
End;

Procedure TOT_GCTYPESTATPIECE.OnDeleteRecord;
var  SQL : string ;
begin
SQL:='SELECT YX_CODE from CHOIXEXT where YX_TYPE="GTL" AND YX_CODE LIKE "'+getField('CC_CODE')+'%"' ;
if ExisteSQL (SQL) then begin LastError:=2; LastErrorMsg:=TexteMessage[LastError]; end ;
End;

Function ExisteFamille ( Code : String ; Niv : integer ) : Boolean ;
Var NomChamp : String ;
BEGIN
Result:=True ;
if Niv>3 then
    BEGIN
    NomChamp:='GA2_FAMILLENIV'+IntToStr(Niv)  ; if ExisteSQL('SELECT GA2_ARTICLE FROM ARTICLECOMPL WHERE '+NomChamp+'="'+Code+'"') then Exit ;
    END else
    BEGIN
    NomChamp:='GA_FAMILLENIV'+IntToStr(Niv)  ; if ExisteSQL('SELECT GA_ARTICLE FROM ARTICLE WHERE '+NomChamp+'="'+Code+'"') then Exit ;
    NomChamp:='GPF_FAMILLENIV'+IntToStr(Niv) ; if ExisteSQL('SELECT GPF_PROFILARTICLE FROM PROFILART WHERE '+NomChamp+'="'+Code+'"') then Exit ;
    NomChamp:='GL_FAMILLENIV'+IntToStr(Niv)  ; if ExisteSQL('SELECT GL_ARTICLE FROM LIGNE WHERE '+NomChamp+'="'+Code+'"') then Exit ;
    END ;
Result:=False ;
END ;

Procedure TOT_GCFAMILLENIV1.OnDeleteRecord;
var  Code : string ;
begin
Code:=GetField('CC_CODE') ;
if ExisteFamille(Code,1) then begin LastError:=3; LastErrorMsg:=TexteMessage[LastError]; end ;
End;

Procedure TOT_GCFAMILLENIV2.OnDeleteRecord;
var  Code : string ;
begin
Code:=GetField('CC_CODE') ;
if ExisteFamille(Code,2) then begin LastError:=3; LastErrorMsg:=TexteMessage[LastError]; end ;
End;

Procedure TOT_GCFAMILLENIV3.OnDeleteRecord;
var  Code : string ;
begin
Code:=GetField('CC_CODE') ;
if ExisteFamille(Code,3) then begin LastError:=3; LastErrorMsg:=TexteMessage[LastError]; end ;
End;

Procedure TOT_GCFAMILLENIV4.OnDeleteRecord;
var  Code : string ;
begin
Code:=GetField('CC_CODE') ;
if ExisteFamille(Code,4) then begin LastError:=3; LastErrorMsg:=TexteMessage[LastError]; end ;
End;

Procedure TOT_GCFAMILLENIV5.OnDeleteRecord;
var  Code : string ;
begin
Code:=GetField('CC_CODE') ;
if ExisteFamille(Code,5) then begin LastError:=3; LastErrorMsg:=TexteMessage[LastError]; end ;
End;

Procedure TOT_GCFAMILLENIV6.OnDeleteRecord;
var  Code : string ;
begin
Code:=GetField('CC_CODE') ;
if ExisteFamille(Code,6) then begin LastError:=3; LastErrorMsg:=TexteMessage[LastError]; end ;
End;

Procedure TOT_GCFAMILLENIV7.OnDeleteRecord;
var  Code : string ;
begin
Code:=GetField('CC_CODE') ;
if ExisteFamille(Code,7) then begin LastError:=3; LastErrorMsg:=TexteMessage[LastError]; end ;
End;

Procedure TOT_GCFAMILLENIV8.OnDeleteRecord;
var  Code : string ;
begin
Code:=GetField('CC_CODE') ;
if ExisteFamille(Code,8) then begin LastError:=3; LastErrorMsg:=TexteMessage[LastError]; end ;
End;

Function ExisteCollection ( Code : String ) : Boolean ;
BEGIN
Result:=True ;
if ExisteSQL('SELECT GA_ARTICLE FROM ARTICLE WHERE GA_COLLECTION="'+Code+'"') then Exit ;
if ExisteSQL('SELECT GL_ARTICLE FROM LIGNE WHERE GL_COLLECTION="'+Code+'"') then Exit ;
Result:=False ;
END ;

Procedure TOT_GCCOLLECTION.OnDeleteRecord;
var  Code : string ;
begin
Code:=GetField('CC_CODE') ;
if ExisteCollection(Code) then begin LastError:=4; LastErrorMsg:=TexteMessage[LastError]; end ;
End;

Function ExisteTarifArt ( Code : String ) : Boolean ;
BEGIN
Result:=True ;
// DCA - FQ MODE 10698
// Les 3 requêtes sont obligatoires même sur la table LIGNE : le tarif d'un article a pu être
// modifié après avoir été saisi dans une pièce.
if ExisteSQL('SELECT GF_TARIF FROM TARIF WHERE GF_TARIFARTICLE="'+Code+'"') then Exit ;
if ExisteSQL('SELECT GA_ARTICLE FROM ARTICLE WHERE GA_TARIFARTICLE="'+Code+'"') then Exit ;
if ExisteSQL('SELECT GL_ARTICLE FROM LIGNE WHERE GL_TARIFARTICLE="'+Code+'"') then Exit ;
Result:=False ;
END ;

Procedure TOT_GCTARIFARTICLE.OnDeleteRecord;
var  Code : string ;
begin
Code:=GetField('CC_CODE') ;
if ExisteTarifArt(Code) then begin LastError:=5; LastErrorMsg:=TexteMessage[LastError]; end ;
End;

Function ExisteTarifCli ( Code : String ) : Boolean ;
BEGIN
Result:=True ;
if ExisteSQL('SELECT T_TIERS FROM TIERS WHERE T_TARIFTIERS="'+Code+'"') then Exit ;
if ExisteSQL('SELECT GF_TARIF FROM TARIF WHERE GF_TARIFTIERS="'+Code+'"') then Exit ;
if ExisteSQL('SELECT GL_ARTICLE FROM LIGNE WHERE GL_TARIFTIERS="'+Code+'"') then Exit ;
Result:=False ;
END ;

Procedure TOT_TTTARIFCLIENT.OnDeleteRecord;
var  Code : string ;
begin
Code:=GetField('CC_CODE') ;
if ExisteTarifCli(Code) then begin LastError:=6; LastErrorMsg:=TexteMessage[LastError]; end ;
End;

Function ExisteComptaArticle ( Code : String ) : Boolean ;
BEGIN
Result:=True ;
if ExisteSQL('SELECT GA_ARTICLE FROM ARTICLE WHERE GA_COMPTAARTICLE="'+Code+'"') then Exit ;
Result:=False ;
END ;

Procedure TOT_GCCOMPTAARTICLE.OnDeleteRecord;
var  Code : string ;
begin
Code:=GetField('CC_CODE') ;
if ExisteComptaArticle(Code) then begin LastError:=7; LastErrorMsg:=TexteMessage[LastError]; end ;
End;

Function ExisteComptaTiers ( Code : String ) : Boolean ;
BEGIN
Result:=True ;
if ExisteSQL('SELECT T_COMPTATIERS FROM TIERS WHERE T_COMPTATIERS="'+Code+'"') then Exit ;
Result:=False ;
END ;

Procedure TOT_GCCOMPTATIERS.OnDeleteRecord;
var  Code : string ;
begin
Code:=GetField('CC_CODE') ;
if ExisteComptaTiers(Code) then begin LastError:=8; LastErrorMsg:=TexteMessage[LastError]; end ;
End;

{ Appelé avec stMasque renseigné OU stTypeMasque renseigné.
  1. stTypeMasque renseigné : créer tous les masques avec ce nouveau type de masque
     bMultiEtab est alors renseigné : le type de masque est-il multi-établissements ?
  2. stMasque renseigné : créer tous les masques pour chaque type de masque défini
                          dans ce cas, DS fournit les CODEDIMx et POSITIONx    }

{$IFDEF EAGLCLIENT}
// AFAIREEAGL
{$ELSE}
procedure CreerMasqueAuto(stTypeMasque : string; bMultiEtab : boolean; stMasque : string; DS : TDataSet);
var QQ : TQuery ;
    iPos,iChamp : integer ;
    bPosOccupe,bAChercher,bOk : boolean ;
    SQL,insMasque,insTypeMasque,stPos : string ;
    position : array[1..6] of string ;
begin
if stTypeMasque=VH_GC.BOTypeMasque_Defaut then exit ;
// Position[] définit l'ordre de recherche d'une position libre.
position[1]:='LI1' ; position[2]:='LI2' ; position[3]:='ON1' ;
position[4]:='CO1' ; position[5]:='CO2' ; position[6]:='' ;
if stMasque<>'' then
    BEGIN
    SQL:='select GMQ_TYPEMASQUE,GMQ_MULTIETAB from TYPEMASQUE ' +
         ' where GMQ_TYPEMASQUE<>"'+VH_GC.BOTypeMasque_Defaut +'" and ' +
         ' GMQ_TYPEMASQUE not in (select GDM_TYPEMASQUE from DIMMASQUE where GDM_MASQUE="'+stMasque +
         '" and GDM_TYPEMASQUE<>"'+VH_GC.BOTypeMasque_Defaut +'")' ;
    bOk:=(DS<>nil) ;
    END else
    BEGIN
    SQL:='select GDM_MASQUE,GDM_LIBELLE,GDM_TYPE1,GDM_TYPE2,GDM_TYPE3,GDM_TYPE4,GDM_TYPE5,' +
         'GDM_POSITION1,GDM_POSITION2,GDM_POSITION3,GDM_POSITION4,GDM_POSITION5,GDM_FERMER' +
         ' from DIMMASQUE' +
         ' where GDM_TYPEMASQUE="'+VH_GC.BOTypeMasque_Defaut +
         '" and GDM_MASQUE not in (select GDM_MASQUE from DIMMASQUE' +
                                  ' where GDM_TYPEMASQUE="'+stTypeMasque+'")' ;
    bOk:=True ;
    END ;
QQ:=OpenSQL(SQL,True) ;
if bOk then
  while not QQ.EOF do
    BEGIN
    if stMasque<>'' then
        BEGIN
        insMasque:=stMasque ;
        insTypeMasque:=QQ.FindField('GMQ_TYPEMASQUE').AsString ;
        bAChercher:=(QQ.FindField('GMQ_MULTIETAB').AsString='X') ;
        END else
        BEGIN
        insMasque:=QQ.FindField('GDM_MASQUE').AsString ;
        insTypeMasque:=stTypeMasque ;
        bAChercher:=bMultiEtab ;
        END ;
    if bAChercher then
        BEGIN
        // Recherche de la position libre à laquelle doit être définie l'établissement
        for iPos:=1 to 5 do
            BEGIN
            bPosOccupe:=False ;
            for iChamp:=1 to 5 do
                BEGIN
                if stMasque<>'' then stPos:=DS.FindField('GDM_POSITION'+IntToStr(iChamp)).AsString
                                else stPos:=QQ.FindField('GDM_POSITION'+IntToStr(iChamp)).AsString ;
                if stPos=position[iPos] then BEGIN bPosOccupe:=True ; break END ;
                END ;
            if not bPosOccupe then break ;
            END ;
        END
        else ipos:=6 ;
    // Si non trouvé, iPos vaut 6 et position[iPos]='' -> nouveau masque identique au masque par défaut.
    // Création du nouveau masque
    SQL:='insert into DIMMASQUE(GDM_MASQUE,GDM_TYPEMASQUE,GDM_LIBELLE,GDM_FERMER,' ;
    for iChamp:=1 to 5 do SQL:=SQL+'GDM_TYPE'+IntToStr(iChamp)+',' ;
    for iChamp:=1 to 5 do SQL:=SQL+'GDM_POSITION'+IntToStr(iChamp)+',' ;
    if stMasque<>'' then
        BEGIN
        SQL:=SQL+'GDM_POSITION6) values ("'+insMasque+'","'+insTypeMasque+
                 '","'+DS.FindField('GDM_LIBELLE').AsString+'","'+DS.FindField('GDM_FERMER').AsString+'","' ;
        for iChamp:=1 to 5 do SQL:=SQL+DS.FindField('GDM_TYPE'+IntToStr(iChamp)).AsString+'","' ;
        for iChamp:=1 to 5 do SQL:=SQL+DS.FindField('GDM_POSITION'+IntToStr(iChamp)).AsString+'","' ;
        END else
        BEGIN
        SQL:=SQL+'GDM_POSITION6) values ("'+insMasque+'","'+insTypeMasque+
                 '","'+QQ.FindField('GDM_LIBELLE').AsString+'","'+QQ.FindField('GDM_FERMER').AsString+'","' ;
        for iChamp:=1 to 5 do SQL:=SQL+QQ.FindField('GDM_TYPE'+IntToStr(iChamp)).AsString+'","' ;
        for iChamp:=1 to 5 do SQL:=SQL+QQ.FindField('GDM_POSITION'+IntToStr(iChamp)).AsString+'","' ;
        END ;
    SQL:=SQL+position[iPos]+'")' ;
    ExecuteSQL(SQL) ;
    QQ.next ;
    END ;
ferme(QQ) ;

end ;
{$ENDIF}
Initialization
registerclasses([TOT_GCLIBREART1]) ; registerclasses([TOT_GCLIBREART2]) ;
registerclasses([TOT_GCLIBREART3]) ; registerclasses([TOT_GCLIBREART4]) ;
registerclasses([TOT_GCLIBREART5]) ; registerclasses([TOT_GCLIBREART6]) ;
registerclasses([TOT_GCLIBREART7]) ; registerclasses([TOT_GCLIBREART8]) ;
registerclasses([TOT_GCLIBREART9]) ; registerclasses([TOT_GCLIBREARTA]) ;
registerclasses([TOT_GCLIBRETIERS1]) ; registerclasses([TOT_GCLIBRETIERS2]) ;
registerclasses([TOT_GCLIBRETIERS3]) ; registerclasses([TOT_GCLIBRETIERS4]) ;
registerclasses([TOT_GCLIBRETIERS5]) ; registerclasses([TOT_GCLIBRETIERS6]) ;
registerclasses([TOT_GCLIBRETIERS7]) ; registerclasses([TOT_GCLIBRETIERS8]) ;
registerclasses([TOT_GCLIBRETIERS9]) ; registerclasses([TOT_GCLIBRETIERSA]) ;
registerclasses([TOT_YYLIBREDEP1]) ; registerclasses([TOT_YYLIBREDEP2]) ;
registerclasses([TOT_YYLIBREDEP3]) ; registerclasses([TOT_YYLIBREDEP4]) ;
registerclasses([TOT_YYLIBREDEP5]) ; registerclasses([TOT_YYLIBREDEP6]) ;
registerclasses([TOT_YYLIBREDEP7]) ; registerclasses([TOT_YYLIBREDEP8]) ;
registerclasses([TOT_YYLIBREDEP9]) ; registerclasses([TOT_YYLIBREDEPA]) ;
{$IFNDEF CCS3}
registerclasses([TOT_YYLIBREET1]) ; registerclasses([TOT_YYLIBREET2]) ;
registerclasses([TOT_YYLIBREET3]) ; registerclasses([TOT_YYLIBREET4]) ;
registerclasses([TOT_YYLIBREET5]) ; registerclasses([TOT_YYLIBREET6]) ;
registerclasses([TOT_YYLIBREET7]) ; registerclasses([TOT_YYLIBREET8]) ;
registerclasses([TOT_YYLIBREET9]) ; registerclasses([TOT_YYLIBREETA]) ;
{$ENDIF}
registerclasses([TOT_YYLIBRECON1]) ; registerclasses([TOT_YYLIBRECON2]) ;
registerclasses([TOT_YYLIBRECON3]) ; registerclasses([TOT_GCZONECOM]) ;
registerclasses([TOT_GCZONELIBRE]) ;
registerclasses([TOT_GCLIBFAMILLE]) ;
registerclasses([TOT_GCGRILLEDIM1]) ; registerclasses([TOT_GCGRILLEDIM2]) ;
registerclasses([TOT_GCGRILLEDIM3]) ; registerclasses([TOT_GCGRILLEDIM4]) ;
registerclasses([TOT_GCGRILLEDIM5]) ;
registerclasses([TOT_GCCATEGORIEDIMS3]) ;
registerclasses([TOT_GCTYPESTATPIECE]) ;
registerclasses([TOT_GCLIBREFOU1,TOT_GCLIBREFOU2,TOT_GCLIBREFOU3]) ;
registerclasses([TOT_GCLIBREPIECE1,TOT_GCLIBREPIECE2,TOT_GCLIBREPIECE3]) ;
registerclasses([TOT_GCFAMILLENIV1]) ; registerclasses([TOT_GCFAMILLENIV2]) ; registerclasses([TOT_GCFAMILLENIV3]) ;
registerclasses([TOT_GCFAMILLENIV4]) ; registerclasses([TOT_GCFAMILLENIV5]) ; registerclasses([TOT_GCFAMILLENIV6]) ;
registerclasses([TOT_GCFAMILLENIV7]) ; registerclasses([TOT_GCFAMILLENIV8]) ;
registerclasses([TOT_GCCOLLECTION]) ;
registerclasses([TOT_GCTARIFARTICLE]) ; registerclasses([TOT_TTTARIFCLIENT]) ;
registerclasses([TOT_GCCOMPTAARTICLE]) ; registerclasses([TOT_GCCOMPTATIERS]) ;
registerclasses([TOT_GCCOLLECTION]) ;
end.
