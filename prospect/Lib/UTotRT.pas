unit UTOTRT;

interface

uses Classes,UTOT,SysUtils,HCtrls,HEnt1,KpmgUtil,UTob,
{$IFDEF EAGLCLIENT}
eTablette,Maineagl,
{$ELSE}
Tablette,DbGrids,hdb,FE_MAin,
    {$IFNDEF DBXPRESS}dbtables{BDE},{$ELSE}uDbxDataSet,{$ENDIF}db,

{$ENDIF}
    Lookup;
Type
    TOT_RTMOTIFS= Class (TOT)
        Procedure OnUpdateRecord ; override ;
        procedure OnDeleteRecord  ; override ;
        End;
    TOT_RTTYPEPERSPECTIVE= Class (TOT)
        procedure OnDeleteRecord  ; override ;
        End;
    TOT_RTRPRLIBPERSPECTIVE1= Class (TOT)
        procedure OnDeleteRecord  ; override ;
        End;
    TOT_RTRPRLIBPERSPECTIVE2= Class (TOT)
        procedure OnDeleteRecord  ; override ;
        End;
    TOT_RTRPRLIBPERSPECTIVE3= Class (TOT)
        procedure OnDeleteRecord  ; override ;
        End;

    TOT_RTIMPORTANCEACT= Class (TOT)
        procedure OnDeleteRecord  ; override ;
        End;
    TOT_RTRPRLIBACTION1= Class (TOT)
        procedure OnDeleteRecord  ; override ;
        End;
    TOT_RTRPRLIBACTION2= Class (TOT)
        procedure OnDeleteRecord  ; override ;
        End;
    TOT_RTRPRLIBACTION3= Class (TOT)
        procedure OnDeleteRecord  ; override ;
        End;
    TOT_RT001LIBTABLE0= Class (TOT)
        procedure OnDeleteRecord  ; override ;
        End;
    TOT_RT001LIBTABLE1= Class (TOT)
        procedure OnDeleteRecord  ; override ;
        End;
    TOT_RT001LIBTABLE2= Class (TOT)
        procedure OnDeleteRecord  ; override ;
        End;
    TOT_RT001LIBTABLE3= Class (TOT)
        procedure OnDeleteRecord  ; override ;
        End;
    TOT_RT001LIBTABLE4= Class (TOT)
        procedure OnDeleteRecord  ; override ;
        End;
    TOT_RT001LIBTABMUL0= Class (TOT)
        procedure OnDeleteRecord  ; override ;
        End;
    TOT_RT001LIBTABMUL1= Class (TOT)
        procedure OnDeleteRecord  ; override ;
        End;
    TOT_RT001LIBTABMUL2= Class (TOT)
        procedure OnDeleteRecord  ; override ;
        End;
    TOT_RT001LIBTABMUL3= Class (TOT)
        procedure OnDeleteRecord  ; override ;
        End;
    TOT_RT001LIBTABMUL4= Class (TOT)
        procedure OnDeleteRecord  ; override ;
        End;

    TOT_RTLIBCHAINAGE1= Class (TOT)
        procedure OnDeleteRecord  ; override ;
        End;
    TOT_RTLIBCHAINAGE2= Class (TOT)
        procedure OnDeleteRecord  ; override ;
        End;
    TOT_RTLIBCHAINAGE3= Class (TOT)
        procedure OnDeleteRecord  ; override ;
        End;

    TOT_RTRSCLIBTABLE0= Class (TOT)
        procedure OnDeleteRecord  ; override ;
        End;
    TOT_RTRSCLIBTABLE1= Class (TOT)
        procedure OnDeleteRecord  ; override ;
        End;
    TOT_RTRSCLIBTABLE2= Class (TOT)
        procedure OnDeleteRecord  ; override ;
        End;
    TOT_RTRSCLIBTABLE3= Class (TOT)
        procedure OnDeleteRecord  ; override ;
        End;
    TOT_RTRSCLIBTABLE4= Class (TOT)
        procedure OnDeleteRecord  ; override ;
        End;
    TOT_RTRSCLIBTABLE5= Class (TOT)
        procedure OnDeleteRecord  ; override ;
        End;
    TOT_RTRSCLIBTABLE6= Class (TOT)
        procedure OnDeleteRecord  ; override ;
        End;
    TOT_RTRSCLIBTABLE7= Class (TOT)
        procedure OnDeleteRecord  ; override ;
        End;
    TOT_RTRSCLIBTABLE8= Class (TOT)
        procedure OnDeleteRecord  ; override ;
        End;
    TOT_RTRSCLIBTABLE9= Class (TOT)
        procedure OnDeleteRecord  ; override ;
        End;
    TOT_RTMOTIFFERMETURE= Class (TOT)
        procedure OnDeleteRecord  ; override ;
        End;

    TOT_RTOBJETOPE= Class (TOT)
        procedure OnDeleteRecord  ; override ;
        End;
    TOT_RT002LIBTABLE0= Class (TOT)
        procedure OnDeleteRecord  ; override ;
        End;
    TOT_RT002LIBTABLE1= Class (TOT)
        procedure OnDeleteRecord  ; override ;
        End;
    TOT_RT002LIBTABLE2= Class (TOT)
        procedure OnDeleteRecord  ; override ;
        End;
    TOT_RT002LIBTABLE3= Class (TOT)
        procedure OnDeleteRecord  ; override ;
        End;
    TOT_RT002LIBTABLE4= Class (TOT)
        procedure OnDeleteRecord  ; override ;
        End;
    TOT_RT002LIBTABMUL0= Class (TOT)
        procedure OnDeleteRecord  ; override ;
        End;
    TOT_RT002LIBTABMUL1= Class (TOT)
        procedure OnDeleteRecord  ; override ;
        End;
    TOT_RT002LIBTABMUL2= Class (TOT)
        procedure OnDeleteRecord  ; override ;
        End;
    TOT_RT002LIBTABMUL3= Class (TOT)
        procedure OnDeleteRecord  ; override ;
        End;
    TOT_RT002LIBTABMUL4= Class (TOT)
        procedure OnDeleteRecord  ; override ;
        End;

    TOT_RTLIBGED1= Class (TOT)
        procedure OnDeleteRecord  ; override ;
        End;
    TOT_RTLIBGED2= Class (TOT)
        procedure OnDeleteRecord  ; override ;
        End;
    TOT_RTLIBGED3= Class (TOT)
        procedure OnDeleteRecord  ; override ;
        End;

    TOT_RTRPRLIBTABLE0= Class (TOT)
        procedure OnDeleteRecord  ; override ;
        End;
    TOT_RTRPRLIBTABLE1= Class (TOT)
        procedure OnDeleteRecord  ; override ;
        End;
    TOT_RTRPRLIBTABLE2= Class (TOT)
        procedure OnDeleteRecord  ; override ;
        End;
    TOT_RTRPRLIBTABLE3= Class (TOT)
        procedure OnDeleteRecord  ; override ;
        End;
    TOT_RTRPRLIBTABLE4= Class (TOT)
        procedure OnDeleteRecord  ; override ;
        End;
    TOT_RTRPRLIBTABLE5= Class (TOT)
        procedure OnDeleteRecord  ; override ;
        End;
    TOT_RTRPRLIBTABLE6= Class (TOT)
        procedure OnDeleteRecord  ; override ;
        End;
    TOT_RTRPRLIBTABLE7= Class (TOT)
        procedure OnDeleteRecord  ; override ;
        End;
    TOT_RTRPRLIBTABLE8= Class (TOT)
        procedure OnDeleteRecord  ; override ;
        End;
    TOT_RTRPRLIBTABLE9= Class (TOT)
        procedure OnDeleteRecord  ; override ;
        End;
    TOT_RTRPRLIBTABLE10= Class (TOT)
        procedure OnDeleteRecord  ; override ;
        End;
    TOT_RTRPRLIBTABLE11= Class (TOT)
        procedure OnDeleteRecord  ; override ;
        End;
    TOT_RTRPRLIBTABLE12= Class (TOT)
        procedure OnDeleteRecord  ; override ;
        End;
    TOT_RTRPRLIBTABLE13= Class (TOT)
        procedure OnDeleteRecord  ; override ;
        End;
    TOT_RTRPRLIBTABLE14= Class (TOT)
        procedure OnDeleteRecord  ; override ;
        End;
    TOT_RTRPRLIBTABLE15= Class (TOT)
        procedure OnDeleteRecord  ; override ;
        End;
    TOT_RTRPRLIBTABLE16= Class (TOT)
        procedure OnDeleteRecord  ; override ;
        End;
    TOT_RTRPRLIBTABLE17= Class (TOT)
        procedure OnDeleteRecord  ; override ;
        End;
    TOT_RTRPRLIBTABLE18= Class (TOT)
        procedure OnDeleteRecord  ; override ;
        End;
    TOT_RTRPRLIBTABLE19= Class (TOT)
        procedure OnDeleteRecord  ; override ;
        End;
    TOT_RTRPRLIBTABLE20= Class (TOT)
        procedure OnDeleteRecord  ; override ;
        End;
    TOT_RTRPRLIBTABLE21= Class (TOT)
        procedure OnDeleteRecord  ; override ;
        End;
    TOT_RTRPRLIBTABLE22= Class (TOT)
        procedure OnDeleteRecord  ; override ;
        End;
    TOT_RTRPRLIBTABLE23= Class (TOT)
        procedure OnDeleteRecord  ; override ;
        End;
    TOT_RTRPRLIBTABLE24= Class (TOT)
        procedure OnDeleteRecord  ; override ;
        End;
    TOT_RTRPRLIBTABLE25= Class (TOT)
        procedure OnDeleteRecord  ; override ;
        End;
    TOT_RTRPRLIBTABLE26= Class (TOT)
        procedure OnDeleteRecord  ; override ;
        End;
    TOT_RTRPRLIBTABLE27= Class (TOT)
        procedure OnDeleteRecord  ; override ;
        End;
    TOT_RTRPRLIBTABLE28= Class (TOT)
        procedure OnDeleteRecord  ; override ;
        End;
    TOT_RTRPRLIBTABLE29= Class (TOT)
        procedure OnDeleteRecord  ; override ;
        End;
    TOT_RTRPRLIBTABLE30= Class (TOT)
        procedure OnDeleteRecord  ; override ;
        End;
    TOT_RTRPRLIBTABLE31= Class (TOT)
        procedure OnDeleteRecord  ; override ;
        End;
    TOT_RTRPRLIBTABLE32= Class (TOT)
        procedure OnDeleteRecord  ; override ;
        End;
    TOT_RTRPRLIBTABLE33= Class (TOT)
        procedure OnDeleteRecord  ; override ;
        End;
    TOT_RTRPRLIBTABLE34= Class (TOT)
        procedure OnDeleteRecord  ; override ;
        End;
    TOT_RTRPRLIBTABMUL0= Class (TOT)
        procedure OnDeleteRecord  ; override ;
        End;
    TOT_RTRPRLIBTABMUL1= Class (TOT)
        procedure OnDeleteRecord  ; override ;
        End;
    TOT_RTRPRLIBTABMUL2= Class (TOT)
        procedure OnDeleteRecord  ; override ;
        End;
    TOT_RTRPRLIBTABMUL3= Class (TOT)
        procedure OnDeleteRecord  ; override ;
        End;
    TOT_RTRPRLIBTABMUL4= Class (TOT)
        procedure OnDeleteRecord  ; override ;
        End;
    TOT_RTRPRLIBTABMUL5= Class (TOT)
        procedure OnDeleteRecord  ; override ;
        End;
    TOT_RTRPRLIBTABMUL6= Class (TOT)
        procedure OnDeleteRecord  ; override ;
        End;
    TOT_RTRPRLIBTABMUL7= Class (TOT)
        procedure OnDeleteRecord  ; override ;
        End;
    TOT_RTRPRLIBTABMUL8= Class (TOT)
        procedure OnDeleteRecord  ; override ;
        End;
    TOT_RTRPRLIBTABMUL9= Class (TOT)
        procedure OnDeleteRecord  ; override ;
        End;

    TOT_RTRPCLIBTABLE0= Class (TOT)
        procedure OnDeleteRecord  ; override ;
        End;
    TOT_RTRPCLIBTABLE1= Class (TOT)
        procedure OnDeleteRecord  ; override ;
        End;
    TOT_RTRPCLIBTABLE2= Class (TOT)
        procedure OnDeleteRecord  ; override ;
        End;
    TOT_RTRPCLIBTABLE3= Class (TOT)
        procedure OnDeleteRecord  ; override ;
        End;
    TOT_RTRPCLIBTABLE4= Class (TOT)
        procedure OnDeleteRecord  ; override ;
        End;
    TOT_RTRPCLIBTABLE5= Class (TOT)
        procedure OnDeleteRecord  ; override ;
        End;
    TOT_RTRPCLIBTABLE6= Class (TOT)
        procedure OnDeleteRecord  ; override ;
        End;
    TOT_RTRPCLIBTABLE7= Class (TOT)
        procedure OnDeleteRecord  ; override ;
        End;
    TOT_RTRPCLIBTABLE8= Class (TOT)
        procedure OnDeleteRecord  ; override ;
        End;
    TOT_RTRPCLIBTABLE9= Class (TOT)
        procedure OnDeleteRecord  ; override ;
        End;
    TOT_RTRPCLIBTABLEA= Class (TOT)
        procedure OnDeleteRecord  ; override ;
        End;
    TOT_RTRPCLIBTABLEB= Class (TOT)
        procedure OnDeleteRecord  ; override ;
        End;

    TOT_RFRPRLIBACTION1= Class (TOT)
        procedure OnDeleteRecord  ; override ;
        End;
    TOT_RFRPRLIBACTION2= Class (TOT)
        procedure OnDeleteRecord  ; override ;
        End;
    TOT_RFRPRLIBACTION3= Class (TOT)
        procedure OnDeleteRecord  ; override ;
        End;

    TOT_RFLIBCHAINAGE1= Class (TOT)
        procedure OnDeleteRecord  ; override ;
        End;
    TOT_RFLIBCHAINAGE2= Class (TOT)
        procedure OnDeleteRecord  ; override ;
        End;
    TOT_RFLIBCHAINAGE3= Class (TOT)
        procedure OnDeleteRecord  ; override ;
        End;

    TOT_RFOBJETOPE= Class (TOT)
        procedure OnDeleteRecord  ; override ;
        End;

    TOT_RT003LIBTABLE0= Class (TOT)
        procedure OnDeleteRecord  ; override ;
        End;
    TOT_RT003LIBTABLE1= Class (TOT)
        procedure OnDeleteRecord  ; override ;
        End;
    TOT_RT003LIBTABLE2= Class (TOT)
        procedure OnDeleteRecord  ; override ;
        End;
    TOT_RT003LIBTABLE3= Class (TOT)
        procedure OnDeleteRecord  ; override ;
        End;
    TOT_RT003LIBTABLE4= Class (TOT)
        procedure OnDeleteRecord  ; override ;
        End;

    TOT_RT003LIBTABMUL0= Class (TOT)
        procedure OnDeleteRecord  ; override ;
        End;
    TOT_RT003LIBTABMUL1= Class (TOT)
        procedure OnDeleteRecord  ; override ;
        End;
    TOT_RT003LIBTABMUL2= Class (TOT)
        procedure OnDeleteRecord  ; override ;
        End;
    TOT_RT003LIBTABMUL3= Class (TOT)
        procedure OnDeleteRecord  ; override ;
        End;
    TOT_RT003LIBTABMUL4= Class (TOT)
        procedure OnDeleteRecord  ; override ;
        End;

    TOT_RT006LIBTABLE0= Class (TOT)
        procedure OnDeleteRecord  ; override ;
        End;
    TOT_RT006LIBTABLE1= Class (TOT)
        procedure OnDeleteRecord  ; override ;
        End;
    TOT_RT006LIBTABLE2= Class (TOT)
        procedure OnDeleteRecord  ; override ;
        End;
    TOT_RT006LIBTABLE3= Class (TOT)
        procedure OnDeleteRecord  ; override ;
        End;
    TOT_RT006LIBTABLE4= Class (TOT)
        procedure OnDeleteRecord  ; override ;
        End;

    TOT_RT006LIBTABMUL0= Class (TOT)
        procedure OnDeleteRecord  ; override ;
        End;
    TOT_RT006LIBTABMUL1= Class (TOT)
        procedure OnDeleteRecord  ; override ;
        End;
    TOT_RT006LIBTABMUL2= Class (TOT)
        procedure OnDeleteRecord  ; override ;
        End;
    TOT_RT006LIBTABMUL3= Class (TOT)
        procedure OnDeleteRecord  ; override ;
        End;
    TOT_RT006LIBTABMUL4= Class (TOT)
        procedure OnDeleteRecord  ; override ;
        End;
    TOT_RT00QLIBTABLE0= Class (TOT)
        procedure OnDeleteRecord  ; override ;
        End;
    TOT_RT00QLIBTABLE1= Class (TOT)
        procedure OnDeleteRecord  ; override ;
        End;
    TOT_RT00QLIBTABLE2= Class (TOT)
        procedure OnDeleteRecord  ; override ;
        End;
    TOT_RT00QLIBTABLE3= Class (TOT)
        procedure OnDeleteRecord  ; override ;
        End;
    TOT_RT00QLIBTABLE4= Class (TOT)
        procedure OnDeleteRecord  ; override ;
        End;
    TOT_RT00QLIBTABMUL0= Class (TOT)
        procedure OnDeleteRecord  ; override ;
        End;
    TOT_RT00QLIBTABMUL1= Class (TOT)
        procedure OnDeleteRecord  ; override ;
        End;
    TOT_RT00QLIBTABMUL2= Class (TOT)
        procedure OnDeleteRecord  ; override ;
        End;
    TOT_RT00QLIBTABMUL3= Class (TOT)
        procedure OnDeleteRecord  ; override ;
        End;
    TOT_RT00QLIBTABMUL4= Class (TOT)
        procedure OnDeleteRecord  ; override ;
        End;

    TOT_RT00VLIBTABLE0= Class (TOT)
        procedure OnDeleteRecord  ; override ;
        End;
    TOT_RT00VLIBTABLE1= Class (TOT)
        procedure OnDeleteRecord  ; override ;
        End;
    TOT_RT00VLIBTABLE2= Class (TOT)
        procedure OnDeleteRecord  ; override ;
        End;
    TOT_RT00VLIBTABLE3= Class (TOT)
        procedure OnDeleteRecord  ; override ;
        End;
    TOT_RT00VLIBTABLE4= Class (TOT)
        procedure OnDeleteRecord  ; override ;
        End;
    TOT_RT00VLIBTABMUL0= Class (TOT)
        procedure OnDeleteRecord  ; override ;
        End;
    TOT_RT00VLIBTABMUL1= Class (TOT)
        procedure OnDeleteRecord  ; override ;
        End;
    TOT_RT00VLIBTABMUL2= Class (TOT)
        procedure OnDeleteRecord  ; override ;
        End;
    TOT_RT00VLIBTABMUL3= Class (TOT)
        procedure OnDeleteRecord  ; override ;
        End;
    TOT_RT00VLIBTABMUL4= Class (TOT)
        procedure OnDeleteRecord  ; override ;
        End;

{$IFDEF SAV}
    TOT_WLIBREWIV1= Class (TOT) Procedure OnDeleteRecord ; override ; End;
    TOT_WLIBREWIV2= Class (TOT) Procedure OnDeleteRecord ; override ; End;
    TOT_WLIBREWIV3= Class (TOT) Procedure OnDeleteRecord ; override ; End;
    TOT_WLIBREWIV4= Class (TOT) Procedure OnDeleteRecord ; override ; End;
    TOT_WLIBREWIV5= Class (TOT) Procedure OnDeleteRecord ; override ; End;
    TOT_WLIBREWIV6= Class (TOT) Procedure OnDeleteRecord ; override ; End;
    TOT_WLIBREWIV7= Class (TOT) Procedure OnDeleteRecord ; override ; End;
    TOT_WLIBREWIV8= Class (TOT) Procedure OnDeleteRecord ; override ; End;
    TOT_WLIBREWIV9= Class (TOT) Procedure OnDeleteRecord ; override ; End;
    TOT_WLIBREWIVA= Class (TOT) Procedure OnDeleteRecord ; override ; End;
    TOT_RT007LIBTABLE0= Class (TOT) Procedure OnDeleteRecord ; override ; End;
    TOT_RT007LIBTABLE1= Class (TOT) Procedure OnDeleteRecord ; override ; End;
    TOT_RT007LIBTABLE2= Class (TOT) Procedure OnDeleteRecord ; override ; End;
    TOT_RT007LIBTABLE3= Class (TOT) Procedure OnDeleteRecord ; override ; End;
    TOT_RT007LIBTABLE4= Class (TOT) Procedure OnDeleteRecord ; override ; End;
    TOT_RT007LIBTABMUL0= Class (TOT) Procedure OnDeleteRecord ; override ; End;
    TOT_RT007LIBTABMUL1= Class (TOT) Procedure OnDeleteRecord ; override ; End;
    TOT_RT007LIBTABMUL2= Class (TOT) Procedure OnDeleteRecord ; override ; End;
    TOT_RT007LIBTABMUL3= Class (TOT) Procedure OnDeleteRecord ; override ; End;
    TOT_RT007LIBTABMUL4= Class (TOT) Procedure OnDeleteRecord ; override ; End;
{$ENDIF SAV}

    TOT_RTCORRESPIMPORT = Class (TOT)
        procedure OnArgument ( S : String) ; override;
        procedure FlisteOnElipsisClick ( sender : Tobject) ;
        procedure FListeOnsorted ( Sender : Tobject );
        procedure OnDeleteRecord  ; override ;
        procedure bComplementClick ( Sender: TObject );
        end;
    TOT_RBCATEGORIE= Class (TOT)
        procedure OnDeleteRecord  ; override ;
        End;
    TOT_RBMULTICHOIXBC1= Class (TOT)
        procedure OnDeleteRecord  ; override ;
        End;
    TOT_RBTYPEINFORMATION= Class (TOT)
        procedure OnDeleteRecord  ; override ;
        End;

const TexteMessage: array[1..2] of string = (
        {1}   'Le libellé abrégé doit être S (Signature),P (Perte),E (Etat d''avancement) ou A (Abandon) '
        {2}  ,'Ce code est utilisé dans une table, il ne peut pas être supprimé !'
);

implementation

// ******* TABLE PERSPECTIVES *******
Procedure TOT_RTMOTIFS.OnUpdateRecord;
Begin
//if not(DS.State in [dsInsert,dsEdit]) then DS.edit;
SetField ('CC_ABREGE',Uppercase (GetField ('CC_ABREGE')));
if (GetField ('CC_ABREGE') <> 'P') and (GetField ('CC_ABREGE') <> 'S') and
   (GetField ('CC_ABREGE') <> 'E') and (GetField ('CC_ABREGE') <> 'A') then
   begin
   LastError:=1;
   LastErrorMsg:=TexteMessage[LastError];
   end;
End;

Procedure TOT_RTMOTIFS.OnDeleteRecord;
var  SQL : string ;
Begin
// Test de l'existence du code dans les tables
SQL:='SELECT RPE_MOTIFSIGNATURE from PERSPECTIVES where RPE_MOTIFSIGNATURE="'+getField('CC_CODE')+'"' ;
if ExisteSQL (SQL) then begin LastError:=2; LastErrorMsg:=TexteMessage[LastError]; exit; end ;
SQL:='SELECT RPE_MOTIFPERTE1 from PERSPECTIVES where RPE_MOTIFPERTE1="'+getField('CC_CODE')+'"' ;
if ExisteSQL (SQL) then begin LastError:=2; LastErrorMsg:=TexteMessage[LastError]; exit; end ;
SQL:='SELECT RPE_MOTIFPERTE2 from PERSPECTIVES where RPE_MOTIFPERTE2="'+getField('CC_CODE')+'"' ;
if ExisteSQL (SQL) then begin LastError:=2; LastErrorMsg:=TexteMessage[LastError]; exit; end ;
SQL:='SELECT RPH_MOTIFSIGNATURE from PERSPHISTO where RPH_MOTIFSIGNATURE="'+getField('CC_CODE')+'"' ;
if ExisteSQL (SQL) then begin LastError:=2; LastErrorMsg:=TexteMessage[LastError]; exit; end ;
SQL:='SELECT RPH_MOTIFPERTE1 from PERSPHISTO where RPH_MOTIFPERTE1="'+getField('CC_CODE')+'"' ;
if ExisteSQL (SQL) then begin LastError:=2; LastErrorMsg:=TexteMessage[LastError]; exit; end ;
SQL:='SELECT RPH_MOTIFPERTE2 from PERSPHISTO where RPH_MOTIFPERTE2="'+getField('CC_CODE')+'"' ;
if ExisteSQL (SQL) then begin LastError:=2; LastErrorMsg:=TexteMessage[LastError]; exit; end ;
End;

Procedure TOT_RTTYPEPERSPECTIVE.OnDeleteRecord;
var  SQL : string ;
Begin
// Test de l'existence du code dans les tables
SQL:='SELECT RPE_TYPEPERSPECTIV from PERSPECTIVES where RPE_TYPEPERSPECTIV="'+getField('CC_CODE')+'"' ;
if ExisteSQL (SQL) then begin LastError:=2; LastErrorMsg:=TexteMessage[LastError]; exit; end ;
SQL:='SELECT RPH_TYPEPERSPECTIV from PERSPHISTO where RPH_TYPEPERSPECTIV="'+getField('CC_CODE')+'"' ;
if ExisteSQL (SQL) then begin LastError:=2; LastErrorMsg:=TexteMessage[LastError]; exit; end ;
End;

Procedure TOT_RTRPRLIBPERSPECTIVE1.OnDeleteRecord;
var  SQL : string ;
Begin
// Test de l'existence du code dans les tables
SQL:='SELECT RPE_TABLELIBREPER1 from PERSPECTIVES where RPE_TABLELIBREPER1="'+getField('CC_CODE')+'"' ;
if ExisteSQL (SQL) then begin LastError:=2; LastErrorMsg:=TexteMessage[LastError]; exit; end ;
SQL:='SELECT RPH_TABLELIBREPER1 from PERSPHISTO where RPH_TABLELIBREPER1="'+getField('CC_CODE')+'"' ;
if ExisteSQL (SQL) then begin LastError:=2; LastErrorMsg:=TexteMessage[LastError]; exit; end ;
End;

Procedure TOT_RTRPRLIBPERSPECTIVE2.OnDeleteRecord;
var  SQL : string ;
Begin
// Test de l'existence du code dans les tables
SQL:='SELECT RPE_TABLELIBREPER2 from PERSPECTIVES where RPE_TABLELIBREPER2="'+getField('CC_CODE')+'"' ;
if ExisteSQL (SQL) then begin LastError:=2; LastErrorMsg:=TexteMessage[LastError]; exit; end ;
SQL:='SELECT RPH_TABLELIBREPER2 from PERSPHISTO where RPH_TABLELIBREPER2="'+getField('CC_CODE')+'"' ;
if ExisteSQL (SQL) then begin LastError:=2; LastErrorMsg:=TexteMessage[LastError]; exit; end ;
End;

Procedure TOT_RTRPRLIBPERSPECTIVE3.OnDeleteRecord;
var  SQL : string ;
Begin
// Test de l'existence du code dans les tables
SQL:='SELECT RPE_TABLELIBREPER3 from PERSPECTIVES where RPE_TABLELIBREPER3="'+getField('CC_CODE')+'"' ;
if ExisteSQL (SQL) then begin LastError:=2; LastErrorMsg:=TexteMessage[LastError]; exit; end ;
SQL:='SELECT RPH_TABLELIBREPER3 from PERSPHISTO where RPH_TABLELIBREPER3="'+getField('CC_CODE')+'"' ;
if ExisteSQL (SQL) then begin LastError:=2; LastErrorMsg:=TexteMessage[LastError]; exit; end ;
End;

// ******* TABLE ACTIONS *******
Procedure TOT_RTIMPORTANCEACT.OnDeleteRecord;
var  SQL : string ;
Begin
// Test de l'existence du code dans les tables
SQL:='SELECT RAC_NIVIMP from ACTIONS where RAC_NIVIMP="'+getField('CC_CODE')+'"' ;
if ExisteSQL (SQL) then begin LastError:=2; LastErrorMsg:=TexteMessage[LastError]; exit; end ;
SQL:='SELECT RPA_NIVIMP from PARACTIONS where RPA_NIVIMP="'+getField('CC_CODE')+'"' ;
if ExisteSQL (SQL) then begin LastError:=2; LastErrorMsg:=TexteMessage[LastError]; exit; end ;
End;

Procedure TOT_RTRPRLIBACTION1.OnDeleteRecord;
var  SQL : string ;
Begin
// Test de l'existence du code dans les tables
SQL:='SELECT RAC_TABLELIBRE1 from ACTIONS where RAC_TABLELIBRE1="'+getField('CC_CODE')+'"' ;
if ExisteSQL (SQL) then begin LastError:=2; LastErrorMsg:=TexteMessage[LastError]; exit; end ;
SQL:='SELECT RPA_TABLELIBRE1 from PARACTIONS where RPA_TABLELIBRE1="'+getField('CC_CODE')+'"' ;
if ExisteSQL (SQL) then begin LastError:=2; LastErrorMsg:=TexteMessage[LastError]; exit; end ;
SQL:='SELECT RAG_TABLELIBRE1 from ACTIONSGENERIQUES where RAG_TABLELIBRE1="'+getField('CC_CODE')+'"' ;
if ExisteSQL (SQL) then begin LastError:=2; LastErrorMsg:=TexteMessage[LastError]; exit; end ;
End;

Procedure TOT_RTRPRLIBACTION2.OnDeleteRecord;
var  SQL : string ;
Begin
// Test de l'existence du code dans les tables
SQL:='SELECT RAC_TABLELIBRE2 from ACTIONS where RAC_TABLELIBRE2="'+getField('CC_CODE')+'"' ;
if ExisteSQL (SQL) then begin LastError:=2; LastErrorMsg:=TexteMessage[LastError]; exit; end ;
SQL:='SELECT RPA_TABLELIBRE2 from PARACTIONS where RPA_TABLELIBRE2="'+getField('CC_CODE')+'"' ;
if ExisteSQL (SQL) then begin LastError:=2; LastErrorMsg:=TexteMessage[LastError]; exit; end ;
SQL:='SELECT RAG_TABLELIBRE2 from ACTIONSGENERIQUES where RAG_TABLELIBRE2="'+getField('CC_CODE')+'"' ;
if ExisteSQL (SQL) then begin LastError:=2; LastErrorMsg:=TexteMessage[LastError]; exit; end ;
End;

Procedure TOT_RTRPRLIBACTION3.OnDeleteRecord;
var  SQL : string ;
Begin
// Test de l'existence du code dans les tables
SQL:='SELECT RAC_TABLELIBRE3 from ACTIONS where RAC_TABLELIBRE3="'+getField('CC_CODE')+'"' ;
if ExisteSQL (SQL) then begin LastError:=2; LastErrorMsg:=TexteMessage[LastError]; exit; end ;
SQL:='SELECT RPA_TABLELIBRE3 from PARACTIONS where RPA_TABLELIBRE3="'+getField('CC_CODE')+'"' ;
if ExisteSQL (SQL) then begin LastError:=2; LastErrorMsg:=TexteMessage[LastError]; exit; end ;
SQL:='SELECT RAG_TABLELIBRE3 from ACTIONSGENERIQUES where RAG_TABLELIBRE3="'+getField('CC_CODE')+'"' ;
if ExisteSQL (SQL) then begin LastError:=2; LastErrorMsg:=TexteMessage[LastError]; exit; end ;
End;

Procedure TOT_RT001LIBTABLE0.OnDeleteRecord;
var  SQL : string ;
Begin
// Test de l'existence du code dans les tables
SQL:='SELECT RD1_RD1LIBTABLE0 from RTINFOS001 where RD1_RD1LIBTABLE0="'+getField('CC_CODE')+'"' ;
if ExisteSQL (SQL) then begin LastError:=2; LastErrorMsg:=TexteMessage[LastError]; exit; end ;
End;

Procedure TOT_RT001LIBTABLE1.OnDeleteRecord;
var  SQL : string ;
Begin
// Test de l'existence du code dans les tables
SQL:='SELECT RD1_RD1LIBTABLE1 from RTINFOS001 where RD1_RD1LIBTABLE1="'+getField('CC_CODE')+'"' ;
if ExisteSQL (SQL) then begin LastError:=2; LastErrorMsg:=TexteMessage[LastError]; exit; end ;
End;

Procedure TOT_RT001LIBTABLE2.OnDeleteRecord;
var  SQL : string ;
Begin
// Test de l'existence du code dans les tables
SQL:='SELECT RD1_RD1LIBTABLE2 from RTINFOS001 where RD1_RD1LIBTABLE2="'+getField('CC_CODE')+'"' ;
if ExisteSQL (SQL) then begin LastError:=2; LastErrorMsg:=TexteMessage[LastError]; exit; end ;
End;

Procedure TOT_RT001LIBTABLE3.OnDeleteRecord;
var  SQL : string ;
Begin
// Test de l'existence du code dans les tables
SQL:='SELECT RD1_RD1LIBTABLE3 from RTINFOS001 where RD1_RD1LIBTABLE3="'+getField('CC_CODE')+'"' ;
if ExisteSQL (SQL) then begin LastError:=2; LastErrorMsg:=TexteMessage[LastError]; exit; end ;
End;

Procedure TOT_RT001LIBTABLE4.OnDeleteRecord;
var  SQL : string ;
Begin
// Test de l'existence du code dans les tables
SQL:='SELECT RD1_RD1LIBTABLE4 from RTINFOS001 where RD1_RD1LIBTABLE4="'+getField('CC_CODE')+'"' ;
if ExisteSQL (SQL) then begin LastError:=2; LastErrorMsg:=TexteMessage[LastError]; exit; end ;
End;

Procedure TOT_RT001LIBTABMUL0.OnDeleteRecord;
var  SQL : string ;
Begin
// Test de l'existence du code dans les tables
SQL:='SELECT RD1_RD1LIBMUL0 from RTINFOS001 where RD1_RD1LIBMUL0 like "'+getField('CC_CODE')+';%"' ;
if ExisteSQL (SQL) then begin LastError:=2; LastErrorMsg:=TexteMessage[LastError]; exit; end ;
SQL:='SELECT RD1_RD1LIBMUL0 from RTINFOS001 where RD1_RD1LIBMUL0 like "'+'%;'+getField('CC_CODE')+';%"' ;
if ExisteSQL (SQL) then begin LastError:=2; LastErrorMsg:=TexteMessage[LastError]; exit; end ;
End;

Procedure TOT_RT001LIBTABMUL1.OnDeleteRecord;
var  SQL : string ;
Begin
// Test de l'existence du code dans les tables
SQL:='SELECT RD1_RD1LIBMUL1 from RTINFOS001 where RD1_RD1LIBMUL1 like "'+getField('CC_CODE')+';%"' ;
if ExisteSQL (SQL) then begin LastError:=2; LastErrorMsg:=TexteMessage[LastError]; exit; end ;
SQL:='SELECT RD1_RD1LIBMUL1 from RTINFOS001 where RD1_RD1LIBMUL1 like "'+'%;'+getField('CC_CODE')+';%"' ;
if ExisteSQL (SQL) then begin LastError:=2; LastErrorMsg:=TexteMessage[LastError]; exit; end ;
End;

Procedure TOT_RT001LIBTABMUL2.OnDeleteRecord;
var  SQL : string ;
Begin
// Test de l'existence du code dans les tables
SQL:='SELECT RD1_RD1LIBMUL2 from RTINFOS001 where RD1_RD1LIBMUL2 like "'+getField('CC_CODE')+';%"' ;
if ExisteSQL (SQL) then begin LastError:=2; LastErrorMsg:=TexteMessage[LastError]; exit; end ;
SQL:='SELECT RD1_RD1LIBMUL2 from RTINFOS001 where RD1_RD1LIBMUL2 like "'+'%;'+getField('CC_CODE')+';%"' ;
if ExisteSQL (SQL) then begin LastError:=2; LastErrorMsg:=TexteMessage[LastError]; exit; end ;
End;

Procedure TOT_RT001LIBTABMUL3.OnDeleteRecord;
var  SQL : string ;
Begin
// Test de l'existence du code dans les tables
SQL:='SELECT RD1_RD1LIBMUL3 from RTINFOS001 where RD1_RD1LIBMUL3 like "'+getField('CC_CODE')+';%"' ;
if ExisteSQL (SQL) then begin LastError:=2; LastErrorMsg:=TexteMessage[LastError]; exit; end ;
SQL:='SELECT RD1_RD1LIBMUL3 from RTINFOS001 where RD1_RD1LIBMUL3 like "'+'%;'+getField('CC_CODE')+';%"' ;
if ExisteSQL (SQL) then begin LastError:=2; LastErrorMsg:=TexteMessage[LastError]; exit; end ;
End;

Procedure TOT_RT001LIBTABMUL4.OnDeleteRecord;
var  SQL : string ;
Begin
// Test de l'existence du code dans les tables
SQL:='SELECT RD1_RD1LIBMUL4 from RTINFOS001 where RD1_RD1LIBMUL4 like "'+getField('CC_CODE')+';%"' ;
if ExisteSQL (SQL) then begin LastError:=2; LastErrorMsg:=TexteMessage[LastError]; exit; end ;
SQL:='SELECT RD1_RD1LIBMUL4 from RTINFOS001 where RD1_RD1LIBMUL4 like "'+'%;'+getField('CC_CODE')+';%"' ;
if ExisteSQL (SQL) then begin LastError:=2; LastErrorMsg:=TexteMessage[LastError]; exit; end ;
End;

// ******* TABLE CHAINAGES *******
Procedure TOT_RTLIBCHAINAGE1.OnDeleteRecord;
var  SQL : string ;
Begin
// Test de l'existence du code dans les tables
SQL:='SELECT RCH_TABLELIBRECH1 from ACTIONSCHAINEES where RCH_TABLELIBRECH1="'+getField('CC_CODE')+'"' ;
if ExisteSQL (SQL) then begin LastError:=2; LastErrorMsg:=TexteMessage[LastError]; exit; end ;
SQL:='SELECT RPG_TABLELIBRECH1 from PARCHAINAGES where RPG_TABLELIBRECH1="'+getField('CC_CODE')+'"' ;
if ExisteSQL (SQL) then begin LastError:=2; LastErrorMsg:=TexteMessage[LastError]; exit; end ;
End;

Procedure TOT_RTLIBCHAINAGE2.OnDeleteRecord;
var  SQL : string ;
Begin
// Test de l'existence du code dans les tables
SQL:='SELECT RCH_TABLELIBRECH2 from ACTIONSCHAINEES where RCH_TABLELIBRECH2="'+getField('CC_CODE')+'"' ;
if ExisteSQL (SQL) then begin LastError:=2; LastErrorMsg:=TexteMessage[LastError]; exit; end ;
SQL:='SELECT RPG_TABLELIBRECH2 from PARCHAINAGES where RPG_TABLELIBRECH2="'+getField('CC_CODE')+'"' ;
if ExisteSQL (SQL) then begin LastError:=2; LastErrorMsg:=TexteMessage[LastError]; exit; end ;
End;

Procedure TOT_RTLIBCHAINAGE3.OnDeleteRecord;
var  SQL : string ;
Begin
// Test de l'existence du code dans les tables
SQL:='SELECT RCH_TABLELIBRECH3 from ACTIONSCHAINEES where RCH_TABLELIBRECH3="'+getField('CC_CODE')+'"' ;
if ExisteSQL (SQL) then begin LastError:=2; LastErrorMsg:=TexteMessage[LastError]; exit; end ;
SQL:='SELECT RPG_TABLELIBRECH3 from PARCHAINAGES where RPG_TABLELIBRECH3="'+getField('CC_CODE')+'"' ;
if ExisteSQL (SQL) then begin LastError:=2; LastErrorMsg:=TexteMessage[LastError]; exit; end ;
End;

// ******* TABLE SUSPECTS *******
Procedure TOT_RTRSCLIBTABLE0.OnDeleteRecord;
var  SQL : string ;
Begin
// Test de l'existence du code dans les tables
SQL:='SELECT RSC_RSCLIBTABLE0 from SUSPECTSCOMPL where RSC_RSCLIBTABLE0="'+getField('CC_CODE')+'"' ;
if ExisteSQL (SQL) then begin LastError:=2; LastErrorMsg:=TexteMessage[LastError]; exit; end ;
End;

Procedure TOT_RTRSCLIBTABLE1.OnDeleteRecord;
var  SQL : string ;
Begin
// Test de l'existence du code dans les tables
SQL:='SELECT RSC_RSCLIBTABLE1 from SUSPECTSCOMPL where RSC_RSCLIBTABLE1="'+getField('CC_CODE')+'"' ;
if ExisteSQL (SQL) then begin LastError:=2; LastErrorMsg:=TexteMessage[LastError]; exit; end ;
End;

Procedure TOT_RTRSCLIBTABLE2.OnDeleteRecord;
var  SQL : string ;
Begin
// Test de l'existence du code dans les tables
SQL:='SELECT RSC_RSCLIBTABLE2 from SUSPECTSCOMPL where RSC_RSCLIBTABLE2="'+getField('CC_CODE')+'"' ;
if ExisteSQL (SQL) then begin LastError:=2; LastErrorMsg:=TexteMessage[LastError]; exit; end ;
End;

Procedure TOT_RTRSCLIBTABLE3.OnDeleteRecord;
var  SQL : string ;
Begin
// Test de l'existence du code dans les tables
SQL:='SELECT RSC_RSCLIBTABLE03 from SUSPECTSCOMPL where RSC_RSCLIBTABLE3="'+getField('CC_CODE')+'"' ;
if ExisteSQL (SQL) then begin LastError:=2; LastErrorMsg:=TexteMessage[LastError]; exit; end ;
End;

Procedure TOT_RTRSCLIBTABLE4.OnDeleteRecord;
var  SQL : string ;
Begin
// Test de l'existence du code dans les tables
SQL:='SELECT RSC_RSCLIBTABLE4 from SUSPECTSCOMPL where RSC_RSCLIBTABLE4="'+getField('CC_CODE')+'"' ;
if ExisteSQL (SQL) then begin LastError:=2; LastErrorMsg:=TexteMessage[LastError]; exit; end ;
End;

Procedure TOT_RTRSCLIBTABLE5.OnDeleteRecord;
var  SQL : string ;
Begin
// Test de l'existence du code dans les tables
SQL:='SELECT RSC_RSCLIBTABLE5 from SUSPECTSCOMPL where RSC_RSCLIBTABLE5="'+getField('CC_CODE')+'"' ;
if ExisteSQL (SQL) then begin LastError:=2; LastErrorMsg:=TexteMessage[LastError]; exit; end ;
End;

Procedure TOT_RTRSCLIBTABLE6.OnDeleteRecord;
var  SQL : string ;
Begin
// Test de l'existence du code dans les tables
SQL:='SELECT RSC_RSCLIBTABLE6 from SUSPECTSCOMPL where RSC_RSCLIBTABLE6="'+getField('CC_CODE')+'"' ;
if ExisteSQL (SQL) then begin LastError:=2; LastErrorMsg:=TexteMessage[LastError]; exit; end ;
End;

Procedure TOT_RTRSCLIBTABLE7.OnDeleteRecord;
var  SQL : string ;
Begin
// Test de l'existence du code dans les tables
SQL:='SELECT RSC_RSCLIBTABLE7 from SUSPECTSCOMPL where RSC_RSCLIBTABLE7="'+getField('CC_CODE')+'"' ;
if ExisteSQL (SQL) then begin LastError:=2; LastErrorMsg:=TexteMessage[LastError]; exit; end ;
End;

Procedure TOT_RTRSCLIBTABLE8.OnDeleteRecord;
var  SQL : string ;
Begin
// Test de l'existence du code dans les tables
SQL:='SELECT RSC_RSCLIBTABLE8 from SUSPECTSCOMPL where RSC_RSCLIBTABLE8="'+getField('CC_CODE')+'"' ;
if ExisteSQL (SQL) then begin LastError:=2; LastErrorMsg:=TexteMessage[LastError]; exit; end ;
End;

Procedure TOT_RTRSCLIBTABLE9.OnDeleteRecord;
var  SQL : string ;
Begin
// Test de l'existence du code dans les tables
SQL:='SELECT RSC_RSCLIBTABLE9 from SUSPECTSCOMPL where RSC_RSCLIBTABLE9="'+getField('CC_CODE')+'"' ;
if ExisteSQL (SQL) then begin LastError:=2; LastErrorMsg:=TexteMessage[LastError]; exit; end ;
End;

Procedure TOT_RTMOTIFFERMETURE.OnDeleteRecord;
var  SQL : string ;
Begin
// Test de l'existence du code dans les tables
SQL:='SELECT RSU_MOTIFFERME from SUSPECTS where RSU_MOTIFFERME="'+getField('CC_CODE')+'"' ;
if ExisteSQL (SQL) then begin LastError:=2; LastErrorMsg:=TexteMessage[LastError]; exit; end ;
End;

// ******* TABLE OPERATIONS *******
Procedure TOT_RTOBJETOPE.OnDeleteRecord;
var  SQL : string ;
Begin
// Test de l'existence du code dans les tables
SQL:='SELECT ROP_OBJETOPE from OPERATIONS where ROP_OBJETOPE="'+getField('CC_CODE')+'"' ;
if ExisteSQL (SQL) then begin LastError:=2; LastErrorMsg:=TexteMessage[LastError]; exit; end ;
End;

Procedure TOT_RT002LIBTABLE0.OnDeleteRecord;
var  SQL : string ;
Begin
// Test de l'existence du code dans les tables
SQL:='SELECT RD2_RD2LIBTABLE0 from RTINFOS002 where RD2_RD2LIBTABLE0="'+getField('CC_CODE')+'"' ;
if ExisteSQL (SQL) then begin LastError:=2; LastErrorMsg:=TexteMessage[LastError]; exit; end ;
End;

Procedure TOT_RT002LIBTABLE1.OnDeleteRecord;
var  SQL : string ;
Begin
// Test de l'existence du code dans les tables
SQL:='SELECT RD2_RD2LIBTABLE1 from RTINFOS002 where RD2_RD2LIBTABLE1="'+getField('CC_CODE')+'"' ;
if ExisteSQL (SQL) then begin LastError:=2; LastErrorMsg:=TexteMessage[LastError]; exit; end ;
End;

Procedure TOT_RT002LIBTABLE2.OnDeleteRecord;
var  SQL : string ;
Begin
// Test de l'existence du code dans les tables
SQL:='SELECT RD2_RD2LIBTABLE2 from RTINFOS002 where RD2_RD2LIBTABLE2="'+getField('CC_CODE')+'"' ;
if ExisteSQL (SQL) then begin LastError:=2; LastErrorMsg:=TexteMessage[LastError]; exit; end ;
End;

Procedure TOT_RT002LIBTABLE3.OnDeleteRecord;
var  SQL : string ;
Begin
// Test de l'existence du code dans les tables
SQL:='SELECT RD2_RD2LIBTABLE3 from RTINFOS002 where RD2_RD2LIBTABLE3="'+getField('CC_CODE')+'"' ;
if ExisteSQL (SQL) then begin LastError:=2; LastErrorMsg:=TexteMessage[LastError]; exit; end ;
End;

Procedure TOT_RT002LIBTABLE4.OnDeleteRecord;
var  SQL : string ;
Begin
// Test de l'existence du code dans les tables
SQL:='SELECT RD2_RD2LIBTABLE4 from RTINFOS002 where RD2_RD2LIBTABLE4="'+getField('CC_CODE')+'"' ;
if ExisteSQL (SQL) then begin LastError:=2; LastErrorMsg:=TexteMessage[LastError]; exit; end ;
End;

Procedure TOT_RT002LIBTABMUL0.OnDeleteRecord;
var  SQL : string ;
Begin
// Test de l'existence du code dans les tables
SQL:='SELECT RD2_RD2LIBMUL0 from RTINFOS002 where RD2_RD2LIBMUL0 like "'+getField('CC_CODE')+';%"' ;
if ExisteSQL (SQL) then begin LastError:=2; LastErrorMsg:=TexteMessage[LastError]; exit; end ;
SQL:='SELECT RD2_RD2LIBMUL0 from RTINFOS002 where RD2_RD2LIBMUL0 like "'+'%;'+getField('CC_CODE')+';%"' ;
if ExisteSQL (SQL) then begin LastError:=2; LastErrorMsg:=TexteMessage[LastError]; exit; end ;
End;

Procedure TOT_RT002LIBTABMUL1.OnDeleteRecord;
var  SQL : string ;
Begin
// Test de l'existence du code dans les tables
SQL:='SELECT RD2_RD2LIBMUL1 from RTINFOS002 where RD2_RD2LIBMUL1 like "'+getField('CC_CODE')+';%"' ;
if ExisteSQL (SQL) then begin LastError:=2; LastErrorMsg:=TexteMessage[LastError]; exit; end ;
SQL:='SELECT RD2_RD2LIBMUL1 from RTINFOS002 where RD2_RD2LIBMUL1 like "'+'%;'+getField('CC_CODE')+';%"' ;
if ExisteSQL (SQL) then begin LastError:=2; LastErrorMsg:=TexteMessage[LastError]; exit; end ;
End;

Procedure TOT_RT002LIBTABMUL2.OnDeleteRecord;
var  SQL : string ;
Begin
// Test de l'existence du code dans les tables
SQL:='SELECT RD2_RD2LIBMUL2 from RTINFOS002 where RD2_RD2LIBMUL2 like "'+getField('CC_CODE')+';%"' ;
if ExisteSQL (SQL) then begin LastError:=2; LastErrorMsg:=TexteMessage[LastError]; exit; end ;
SQL:='SELECT RD2_RD2LIBMUL2 from RTINFOS002 where RD2_RD2LIBMUL2 like "'+'%;'+getField('CC_CODE')+';%"' ;
if ExisteSQL (SQL) then begin LastError:=2; LastErrorMsg:=TexteMessage[LastError]; exit; end ;
End;

Procedure TOT_RT002LIBTABMUL3.OnDeleteRecord;
var  SQL : string ;
Begin
// Test de l'existence du code dans les tables
SQL:='SELECT RD2_RD2LIBMUL3 from RTINFOS002 where RD2_RD2LIBMUL3 like "'+getField('CC_CODE')+';%"' ;
if ExisteSQL (SQL) then begin LastError:=2; LastErrorMsg:=TexteMessage[LastError]; exit; end ;
SQL:='SELECT RD2_RD2LIBMUL3 from RTINFOS002 where RD2_RD2LIBMUL3 like "'+'%;'+getField('CC_CODE')+';%"' ;
if ExisteSQL (SQL) then begin LastError:=2; LastErrorMsg:=TexteMessage[LastError]; exit; end ;
End;

Procedure TOT_RT002LIBTABMUL4.OnDeleteRecord;
var  SQL : string ;
Begin
// Test de l'existence du code dans les tables
SQL:='SELECT RD2_RD2LIBMUL4 from RTINFOS002 where RD2_RD2LIBMUL4 like "'+getField('CC_CODE')+';%"' ;
if ExisteSQL (SQL) then begin LastError:=2; LastErrorMsg:=TexteMessage[LastError]; exit; end ;
SQL:='SELECT RD2_RD2LIBMUL4 from RTINFOS002 where RD2_RD2LIBMUL4 like "'+'%;'+getField('CC_CODE')+';%"' ;
if ExisteSQL (SQL) then begin LastError:=2; LastErrorMsg:=TexteMessage[LastError]; exit; end ;
End;

// ******* TABLE DOCUMENTS *******
Procedure TOT_RTLIBGED1.OnDeleteRecord;
var  SQL : string ;
Begin
// Test de l'existence du code dans les tables
SQL:='SELECT RTD_TABLELIBREGED1 from RTDOCUMENT where RTD_TABLELIBREGED1="'+getField('CC_CODE')+'"' ;
if ExisteSQL (SQL) then begin LastError:=2; LastErrorMsg:=TexteMessage[LastError]; exit; end ;
End;

Procedure TOT_RTLIBGED2.OnDeleteRecord;
var  SQL : string ;
Begin
// Test de l'existence du code dans les tables
SQL:='SELECT RTD_TABLELIBREGED2 from RTDOCUMENT where RTD_TABLELIBREGED2="'+getField('CC_CODE')+'"' ;
if ExisteSQL (SQL) then begin LastError:=2; LastErrorMsg:=TexteMessage[LastError]; exit; end ;
End;

Procedure TOT_RTLIBGED3.OnDeleteRecord;
var  SQL : string ;
Begin
// Test de l'existence du code dans les tables
SQL:='SELECT RTD_TABLELIBREGED3 from RTDOCUMENT where RTD_TABLELIBREGED3="'+getField('CC_CODE')+'"' ;
if ExisteSQL (SQL) then begin LastError:=2; LastErrorMsg:=TexteMessage[LastError]; exit; end ;
End;

// ******* TABLE PROSPECTS *******
Procedure TOT_RTRPRLIBTABLE0.OnDeleteRecord;
var  SQL : string ;
Begin
// Test de l'existence du code dans les tables
SQL:='SELECT RPR_RPRLIBTABLE0 from PROSPECTS where RPR_RPRLIBTABLE0="'+getField('CC_CODE')+'"' ;
if ExisteSQL (SQL) then begin LastError:=2; LastErrorMsg:=TexteMessage[LastError]; exit; end ;
End;

Procedure TOT_RTRPRLIBTABLE1.OnDeleteRecord;
var  SQL : string ;
Begin
// Test de l'existence du code dans les tables
SQL:='SELECT RPR_RPRLIBTABLE1 from PROSPECTS where RPR_RPRLIBTABLE1="'+getField('CC_CODE')+'"' ;
if ExisteSQL (SQL) then begin LastError:=2; LastErrorMsg:=TexteMessage[LastError]; exit; end ;
End;

Procedure TOT_RTRPRLIBTABLE2.OnDeleteRecord;
var  SQL : string ;
Begin
// Test de l'existence du code dans les tables
SQL:='SELECT RPR_RPRLIBTABLE2 from PROSPECTS where RPR_RPRLIBTABLE2="'+getField('CC_CODE')+'"' ;
if ExisteSQL (SQL) then begin LastError:=2; LastErrorMsg:=TexteMessage[LastError]; exit; end ;
End;

Procedure TOT_RTRPRLIBTABLE3.OnDeleteRecord;
var  SQL : string ;
Begin
// Test de l'existence du code dans les tables
SQL:='SELECT RPR_RPRLIBTABLE3 from PROSPECTS where RPR_RPRLIBTABLE3="'+getField('CC_CODE')+'"' ;
if ExisteSQL (SQL) then begin LastError:=2; LastErrorMsg:=TexteMessage[LastError]; exit; end ;
End;

Procedure TOT_RTRPRLIBTABLE4.OnDeleteRecord;
var  SQL : string ;
Begin
// Test de l'existence du code dans les tables
SQL:='SELECT RPR_RPRLIBTABLE4 from PROSPECTS where RPR_RPRLIBTABLE4="'+getField('CC_CODE')+'"' ;
if ExisteSQL (SQL) then begin LastError:=2; LastErrorMsg:=TexteMessage[LastError]; exit; end ;
End;

Procedure TOT_RTRPRLIBTABLE5.OnDeleteRecord;
var  SQL : string ;
Begin
// Test de l'existence du code dans les tables
SQL:='SELECT RPR_RPRLIBTABLE5 from PROSPECTS where RPR_RPRLIBTABLE5="'+getField('CC_CODE')+'"' ;
if ExisteSQL (SQL) then begin LastError:=2; LastErrorMsg:=TexteMessage[LastError]; exit; end ;
End;

Procedure TOT_RTRPRLIBTABLE6.OnDeleteRecord;
var  SQL : string ;
Begin
// Test de l'existence du code dans les tables
SQL:='SELECT RPR_RPRLIBTABLE6 from PROSPECTS where RPR_RPRLIBTABLE6="'+getField('CC_CODE')+'"' ;
if ExisteSQL (SQL) then begin LastError:=2; LastErrorMsg:=TexteMessage[LastError]; exit; end ;
End;

Procedure TOT_RTRPRLIBTABLE7.OnDeleteRecord;
var  SQL : string ;
Begin
// Test de l'existence du code dans les tables
SQL:='SELECT RPR_RPRLIBTABLE7 from PROSPECTS where RPR_RPRLIBTABLE7="'+getField('CC_CODE')+'"' ;
if ExisteSQL (SQL) then begin LastError:=2; LastErrorMsg:=TexteMessage[LastError]; exit; end ;
End;

Procedure TOT_RTRPRLIBTABLE8.OnDeleteRecord;
var  SQL : string ;
Begin
// Test de l'existence du code dans les tables
SQL:='SELECT RPR_RPRLIBTABLE8 from PROSPECTS where RPR_RPRLIBTABLE8="'+getField('CC_CODE')+'"' ;
if ExisteSQL (SQL) then begin LastError:=2; LastErrorMsg:=TexteMessage[LastError]; exit; end ;
End;

Procedure TOT_RTRPRLIBTABLE9.OnDeleteRecord;
var  SQL : string ;
Begin
// Test de l'existence du code dans les tables
SQL:='SELECT RPR_RPRLIBTABLE9 from PROSPECTS where RPR_RPRLIBTABLE9="'+getField('CC_CODE')+'"' ;
if ExisteSQL (SQL) then begin LastError:=2; LastErrorMsg:=TexteMessage[LastError]; exit; end ;
End;

Procedure TOT_RTRPRLIBTABLE10.OnDeleteRecord;
var  SQL : string ;
Begin
// Test de l'existence du code dans les tables
SQL:='SELECT RPR_RPRLIBTABLE10 from PROSPECTS where RPR_RPRLIBTABLE10="'+getField('CC_CODE')+'"' ;
if ExisteSQL (SQL) then begin LastError:=2; LastErrorMsg:=TexteMessage[LastError]; exit; end ;
End;

Procedure TOT_RTRPRLIBTABLE11.OnDeleteRecord;
var  SQL : string ;
Begin
// Test de l'existence du code dans les tables
SQL:='SELECT RPR_RPRLIBTABLE11 from PROSPECTS where RPR_RPRLIBTABLE11="'+getField('CC_CODE')+'"' ;
if ExisteSQL (SQL) then begin LastError:=2; LastErrorMsg:=TexteMessage[LastError]; exit; end ;
End;

Procedure TOT_RTRPRLIBTABLE12.OnDeleteRecord;
var  SQL : string ;
Begin
// Test de l'existence du code dans les tables
SQL:='SELECT RPR_RPRLIBTABLE12 from PROSPECTS where RPR_RPRLIBTABLE12="'+getField('CC_CODE')+'"' ;
if ExisteSQL (SQL) then begin LastError:=2; LastErrorMsg:=TexteMessage[LastError]; exit; end ;
End;

Procedure TOT_RTRPRLIBTABLE13.OnDeleteRecord;
var  SQL : string ;
Begin
// Test de l'existence du code dans les tables
SQL:='SELECT RPR_RPRLIBTABLE13 from PROSPECTS where RPR_RPRLIBTABLE13="'+getField('CC_CODE')+'"' ;
if ExisteSQL (SQL) then begin LastError:=2; LastErrorMsg:=TexteMessage[LastError]; exit; end ;
End;

Procedure TOT_RTRPRLIBTABLE14.OnDeleteRecord;
var  SQL : string ;
Begin
// Test de l'existence du code dans les tables
SQL:='SELECT RPR_RPRLIBTABLE14 from PROSPECTS where RPR_RPRLIBTABLE14="'+getField('CC_CODE')+'"' ;
if ExisteSQL (SQL) then begin LastError:=2; LastErrorMsg:=TexteMessage[LastError]; exit; end ;
End;

Procedure TOT_RTRPRLIBTABLE15.OnDeleteRecord;
var  SQL : string ;
Begin
// Test de l'existence du code dans les tables
SQL:='SELECT RPR_RPRLIBTABLE15 from PROSPECTS where RPR_RPRLIBTABLE15="'+getField('CC_CODE')+'"' ;
if ExisteSQL (SQL) then begin LastError:=2; LastErrorMsg:=TexteMessage[LastError]; exit; end ;
End;

Procedure TOT_RTRPRLIBTABLE16.OnDeleteRecord;
var  SQL : string ;
Begin
// Test de l'existence du code dans les tables
SQL:='SELECT RPR_RPRLIBTABLE16 from PROSPECTS where RPR_RPRLIBTABLE16="'+getField('CC_CODE')+'"' ;
if ExisteSQL (SQL) then begin LastError:=2; LastErrorMsg:=TexteMessage[LastError]; exit; end ;
End;

Procedure TOT_RTRPRLIBTABLE17.OnDeleteRecord;
var  SQL : string ;
Begin
// Test de l'existence du code dans les tables
SQL:='SELECT RPR_RPRLIBTABLE17 from PROSPECTS where RPR_RPRLIBTABLE17="'+getField('CC_CODE')+'"' ;
if ExisteSQL (SQL) then begin LastError:=2; LastErrorMsg:=TexteMessage[LastError]; exit; end ;
End;

Procedure TOT_RTRPRLIBTABLE18.OnDeleteRecord;
var  SQL : string ;
Begin
// Test de l'existence du code dans les tables
SQL:='SELECT RPR_RPRLIBTABLE18 from PROSPECTS where RPR_RPRLIBTABLE18="'+getField('CC_CODE')+'"' ;
if ExisteSQL (SQL) then begin LastError:=2; LastErrorMsg:=TexteMessage[LastError]; exit; end ;
End;

Procedure TOT_RTRPRLIBTABLE19.OnDeleteRecord;
var  SQL : string ;
Begin
// Test de l'existence du code dans les tables
SQL:='SELECT RPR_RPRLIBTABLE19 from PROSPECTS where RPR_RPRLIBTABLE19="'+getField('CC_CODE')+'"' ;
if ExisteSQL (SQL) then begin LastError:=2; LastErrorMsg:=TexteMessage[LastError]; exit; end ;
End;

Procedure TOT_RTRPRLIBTABLE20.OnDeleteRecord;
var  SQL : string ;
Begin
// Test de l'existence du code dans les tables
SQL:='SELECT RPR_RPRLIBTABLE20 from PROSPECTS where RPR_RPRLIBTABLE20="'+getField('CC_CODE')+'"' ;
if ExisteSQL (SQL) then begin LastError:=2; LastErrorMsg:=TexteMessage[LastError]; exit; end ;
End;

Procedure TOT_RTRPRLIBTABLE21.OnDeleteRecord;
var  SQL : string ;
Begin
// Test de l'existence du code dans les tables
SQL:='SELECT RPR_RPRLIBTABLE21 from PROSPECTS where RPR_RPRLIBTABLE21="'+getField('CC_CODE')+'"' ;
if ExisteSQL (SQL) then begin LastError:=2; LastErrorMsg:=TexteMessage[LastError]; exit; end ;
End;

Procedure TOT_RTRPRLIBTABLE22.OnDeleteRecord;
var  SQL : string ;
Begin
// Test de l'existence du code dans les tables
SQL:='SELECT RPR_RPRLIBTABLE22 from PROSPECTS where RPR_RPRLIBTABLE22="'+getField('CC_CODE')+'"' ;
if ExisteSQL (SQL) then begin LastError:=2; LastErrorMsg:=TexteMessage[LastError]; exit; end ;
End;

Procedure TOT_RTRPRLIBTABLE23.OnDeleteRecord;
var  SQL : string ;
Begin
// Test de l'existence du code dans les tables
SQL:='SELECT RPR_RPRLIBTABLE23 from PROSPECTS where RPR_RPRLIBTABLE23="'+getField('CC_CODE')+'"' ;
if ExisteSQL (SQL) then begin LastError:=2; LastErrorMsg:=TexteMessage[LastError]; exit; end ;
End;

Procedure TOT_RTRPRLIBTABLE24.OnDeleteRecord;
var  SQL : string ;
Begin
// Test de l'existence du code dans les tables
SQL:='SELECT RPR_RPRLIBTABLE24 from PROSPECTS where RPR_RPRLIBTABLE24="'+getField('CC_CODE')+'"' ;
if ExisteSQL (SQL) then begin LastError:=2; LastErrorMsg:=TexteMessage[LastError]; exit; end ;
End;

Procedure TOT_RTRPRLIBTABLE25.OnDeleteRecord;
var  SQL : string ;
Begin
// Test de l'existence du code dans les tables
SQL:='SELECT RPR_RPRLIBTABLE25 from PROSPECTS where RPR_RPRLIBTABLE25="'+getField('CC_CODE')+'"' ;
if ExisteSQL (SQL) then begin LastError:=2; LastErrorMsg:=TexteMessage[LastError]; exit; end ;
End;

Procedure TOT_RTRPRLIBTABLE26.OnDeleteRecord;
var  SQL : string ;
Begin
// Test de l'existence du code dans les tables
SQL:='SELECT RPR_RPRLIBTABLE26 from PROSPECTS where RPR_RPRLIBTABLE26="'+getField('CC_CODE')+'"' ;
if ExisteSQL (SQL) then begin LastError:=2; LastErrorMsg:=TexteMessage[LastError]; exit; end ;
End;

Procedure TOT_RTRPRLIBTABLE27.OnDeleteRecord;
var  SQL : string ;
Begin
// Test de l'existence du code dans les tables
SQL:='SELECT RPR_RPRLIBTABLE27 from PROSPECTS where RPR_RPRLIBTABLE27="'+getField('CC_CODE')+'"' ;
if ExisteSQL (SQL) then begin LastError:=2; LastErrorMsg:=TexteMessage[LastError]; exit; end ;
End;

Procedure TOT_RTRPRLIBTABLE28.OnDeleteRecord;
var  SQL : string ;
Begin
// Test de l'existence du code dans les tables
SQL:='SELECT RPR_RPRLIBTABLE28 from PROSPECTS where RPR_RPRLIBTABLE28="'+getField('CC_CODE')+'"' ;
if ExisteSQL (SQL) then begin LastError:=2; LastErrorMsg:=TexteMessage[LastError]; exit; end ;
End;

Procedure TOT_RTRPRLIBTABLE29.OnDeleteRecord;
var  SQL : string ;
Begin
// Test de l'existence du code dans les tables
SQL:='SELECT RPR_RPRLIBTABLE29 from PROSPECTS where RPR_RPRLIBTABLE29="'+getField('CC_CODE')+'"' ;
if ExisteSQL (SQL) then begin LastError:=2; LastErrorMsg:=TexteMessage[LastError]; exit; end ;
End;

Procedure TOT_RTRPRLIBTABLE30.OnDeleteRecord;
var  SQL : string ;
Begin
// Test de l'existence du code dans les tables
SQL:='SELECT RPR_RPRLIBTABLE30 from PROSPECTS where RPR_RPRLIBTABLE30="'+getField('CC_CODE')+'"' ;
if ExisteSQL (SQL) then begin LastError:=2; LastErrorMsg:=TexteMessage[LastError]; exit; end ;
End;

Procedure TOT_RTRPRLIBTABLE31.OnDeleteRecord;
var  SQL : string ;
Begin
// Test de l'existence du code dans les tables
SQL:='SELECT RPR_RPRLIBTABLE31 from PROSPECTS where RPR_RPRLIBTABLE31="'+getField('CC_CODE')+'"' ;
if ExisteSQL (SQL) then begin LastError:=2; LastErrorMsg:=TexteMessage[LastError]; exit; end ;
End;

Procedure TOT_RTRPRLIBTABLE32.OnDeleteRecord;
var  SQL : string ;
Begin
// Test de l'existence du code dans les tables
SQL:='SELECT RPR_RPRLIBTABLE32 from PROSPECTS where RPR_RPRLIBTABLE32="'+getField('CC_CODE')+'"' ;
if ExisteSQL (SQL) then begin LastError:=2; LastErrorMsg:=TexteMessage[LastError]; exit; end ;
End;

Procedure TOT_RTRPRLIBTABLE33.OnDeleteRecord;
var  SQL : string ;
Begin
// Test de l'existence du code dans les tables
SQL:='SELECT RPR_RPRLIBTABLE33 from PROSPECTS where RPR_RPRLIBTABLE33="'+getField('CC_CODE')+'"' ;
if ExisteSQL (SQL) then begin LastError:=2; LastErrorMsg:=TexteMessage[LastError]; exit; end ;
End;

Procedure TOT_RTRPRLIBTABLE34.OnDeleteRecord;
var  SQL : string ;
Begin
// Test de l'existence du code dans les tables
SQL:='SELECT RPR_RPRLIBTABLE34 from PROSPECTS where RPR_RPRLIBTABLE34="'+getField('CC_CODE')+'"' ;
if ExisteSQL (SQL) then begin LastError:=2; LastErrorMsg:=TexteMessage[LastError]; exit; end ;
End;

Procedure TOT_RTRPRLIBTABMUL0.OnDeleteRecord;
var  SQL : string ;
Begin
// Test de l'existence du code dans les tables
SQL:='SELECT RPR_RPRLIBMUL0 from PROSPECTS where RPR_RPRLIBMUL0 like "'+getField('CC_CODE')+';%"' ;
if ExisteSQL (SQL) then begin LastError:=2; LastErrorMsg:=TexteMessage[LastError]; exit; end ;
SQL:='SELECT RPR_RPRLIBMUL0 from PROSPECTS where RPR_RPRLIBMUL0 like "'+'%;'+getField('CC_CODE')+';%"' ;
if ExisteSQL (SQL) then begin LastError:=2; LastErrorMsg:=TexteMessage[LastError]; exit; end ;
End;

Procedure TOT_RTRPRLIBTABMUL1.OnDeleteRecord;
var  SQL : string ;
Begin
// Test de l'existence du code dans les tables
SQL:='SELECT RPR_RPRLIBMUL1 from PROSPECTS where RPR_RPRLIBMUL1 like "'+getField('CC_CODE')+';%"' ;
if ExisteSQL (SQL) then begin LastError:=2; LastErrorMsg:=TexteMessage[LastError]; exit; end ;
SQL:='SELECT RPR_RPRLIBMUL1 from PROSPECTS where RPR_RPRLIBMUL1 like "'+'%;'+getField('CC_CODE')+';%"' ;
if ExisteSQL (SQL) then begin LastError:=2; LastErrorMsg:=TexteMessage[LastError]; exit; end ;
End;

Procedure TOT_RTRPRLIBTABMUL2.OnDeleteRecord;
var  SQL : string ;
Begin
// Test de l'existence du code dans les tables
SQL:='SELECT RPR_RPRLIBMUL2 from PROSPECTS where RPR_RPRLIBMUL2 like "'+getField('CC_CODE')+';%"' ;
if ExisteSQL (SQL) then begin LastError:=2; LastErrorMsg:=TexteMessage[LastError]; exit; end ;
SQL:='SELECT RPR_RPRLIBMUL2 from PROSPECTS where RPR_RPRLIBMUL2 like "'+'%;'+getField('CC_CODE')+';%"' ;
if ExisteSQL (SQL) then begin LastError:=2; LastErrorMsg:=TexteMessage[LastError]; exit; end ;
End;

Procedure TOT_RTRPRLIBTABMUL3.OnDeleteRecord;
var  SQL : string ;
Begin
// Test de l'existence du code dans les tables
SQL:='SELECT RPR_RPRLIBMUL3 from PROSPECTS where RPR_RPRLIBMUL3 like "'+getField('CC_CODE')+';%"' ;
if ExisteSQL (SQL) then begin LastError:=2; LastErrorMsg:=TexteMessage[LastError]; exit; end ;
SQL:='SELECT RPR_RPRLIBMUL3 from PROSPECTS where RPR_RPRLIBMUL3 like "'+'%;'+getField('CC_CODE')+';%"' ;
if ExisteSQL (SQL) then begin LastError:=2; LastErrorMsg:=TexteMessage[LastError]; exit; end ;
End;

Procedure TOT_RTRPRLIBTABMUL4.OnDeleteRecord;
var  SQL : string ;
Begin
// Test de l'existence du code dans les tables
SQL:='SELECT RPR_RPRLIBMUL4 from PROSPECTS where RPR_RPRLIBMUL4 like "'+getField('CC_CODE')+';%"' ;
if ExisteSQL (SQL) then begin LastError:=2; LastErrorMsg:=TexteMessage[LastError]; exit; end ;
SQL:='SELECT RPR_RPRLIBMUL4 from PROSPECTS where RPR_RPRLIBMUL4 like "'+'%;'+getField('CC_CODE')+';%"' ;
if ExisteSQL (SQL) then begin LastError:=2; LastErrorMsg:=TexteMessage[LastError]; exit; end ;
End;

Procedure TOT_RTRPRLIBTABMUL5.OnDeleteRecord;
var  SQL : string ;
Begin
// Test de l'existence du code dans les tables
SQL:='SELECT RPR_RPRLIBMUL5 from PROSPECTS where RPR_RPRLIBMUL5 like "'+getField('CC_CODE')+';%"' ;
if ExisteSQL (SQL) then begin LastError:=2; LastErrorMsg:=TexteMessage[LastError]; exit; end ;
SQL:='SELECT RPR_RPRLIBMUL5 from PROSPECTS where RPR_RPRLIBMUL5 like "'+'%;'+getField('CC_CODE')+';%"' ;
if ExisteSQL (SQL) then begin LastError:=2; LastErrorMsg:=TexteMessage[LastError]; exit; end ;
End;

Procedure TOT_RTRPRLIBTABMUL6.OnDeleteRecord;
var  SQL : string ;
Begin
// Test de l'existence du code dans les tables
SQL:='SELECT RPR_RPRLIBMUL6 from PROSPECTS where RPR_RPRLIBMUL6 like "'+getField('CC_CODE')+';%"' ;
if ExisteSQL (SQL) then begin LastError:=2; LastErrorMsg:=TexteMessage[LastError]; exit; end ;
SQL:='SELECT RPR_RPRLIBMUL6 from PROSPECTS where RPR_RPRLIBMUL6 like "'+'%;'+getField('CC_CODE')+';%"' ;
if ExisteSQL (SQL) then begin LastError:=2; LastErrorMsg:=TexteMessage[LastError]; exit; end ;
End;

Procedure TOT_RTRPRLIBTABMUL7.OnDeleteRecord;
var  SQL : string ;
Begin
// Test de l'existence du code dans les tables
SQL:='SELECT RPR_RPRLIBMUL7 from PROSPECTS where RPR_RPRLIBMUL7 like "'+getField('CC_CODE')+';%"' ;
if ExisteSQL (SQL) then begin LastError:=2; LastErrorMsg:=TexteMessage[LastError]; exit; end ;
SQL:='SELECT RPR_RPRLIBMUL7 from PROSPECTS where RPR_RPRLIBMUL7 like "'+'%;'+getField('CC_CODE')+';%"' ;
if ExisteSQL (SQL) then begin LastError:=2; LastErrorMsg:=TexteMessage[LastError]; exit; end ;
End;

Procedure TOT_RTRPRLIBTABMUL8.OnDeleteRecord;
var  SQL : string ;
Begin
// Test de l'existence du code dans les tables
SQL:='SELECT RPR_RPRLIBMUL8 from PROSPECTS where RPR_RPRLIBMUL8 like "'+getField('CC_CODE')+';%"' ;
if ExisteSQL (SQL) then begin LastError:=2; LastErrorMsg:=TexteMessage[LastError]; exit; end ;
SQL:='SELECT RPR_RPRLIBMUL8 from PROSPECTS where RPR_RPRLIBMUL8 like "'+'%;'+getField('CC_CODE')+';%"' ;
if ExisteSQL (SQL) then begin LastError:=2; LastErrorMsg:=TexteMessage[LastError]; exit; end ;
End;

Procedure TOT_RTRPRLIBTABMUL9.OnDeleteRecord;
var  SQL : string ;
Begin
// Test de l'existence du code dans les tables
SQL:='SELECT RPR_RPRLIBMUL9 from PROSPECTS where RPR_RPRLIBMUL9 like "'+getField('CC_CODE')+';%"' ;
if ExisteSQL (SQL) then begin LastError:=2; LastErrorMsg:=TexteMessage[LastError]; exit; end ;
SQL:='SELECT RPR_RPRLIBMUL9 from PROSPECTS where RPR_RPRLIBMUL9 like "'+'%;'+getField('CC_CODE')+';%"' ;
if ExisteSQL (SQL) then begin LastError:=2; LastErrorMsg:=TexteMessage[LastError]; exit; end ;
End;

// ******* TABLE COMPLEMENTAIRE PROSPECTS *******
Procedure TOT_RTRPCLIBTABLE0.OnDeleteRecord;
var  SQL : string ;
Begin
// Test de l'existence du code dans les tables
SQL:='SELECT RPC_RPCLIBTABLE0 from PROSPECTCOMPL where RPC_RPCLIBTABLE0="'+getField('CC_CODE')+'"' ;
if ExisteSQL (SQL) then begin LastError:=2; LastErrorMsg:=TexteMessage[LastError]; exit; end ;
End;

Procedure TOT_RTRPCLIBTABLE1.OnDeleteRecord;
var  SQL : string ;
Begin
// Test de l'existence du code dans les tables
SQL:='SELECT RPC_RPCLIBTABLE1 from PROSPECTCOMPL where RPC_RPCLIBTABLE1="'+getField('CC_CODE')+'"' ;
if ExisteSQL (SQL) then begin LastError:=2; LastErrorMsg:=TexteMessage[LastError]; exit; end ;
End;

Procedure TOT_RTRPCLIBTABLE2.OnDeleteRecord;
var  SQL : string ;
Begin
// Test de l'existence du code dans les tables
SQL:='SELECT RPC_RPCLIBTABLE2 from PROSPECTCOMPL where RPC_RPCLIBTABLE2="'+getField('CC_CODE')+'"' ;
if ExisteSQL (SQL) then begin LastError:=2; LastErrorMsg:=TexteMessage[LastError]; exit; end ;
End;

Procedure TOT_RTRPCLIBTABLE3.OnDeleteRecord;
var  SQL : string ;
Begin
// Test de l'existence du code dans les tables
SQL:='SELECT RPC_RPCLIBTABLE3 from PROSPECTCOMPL where RPC_RPCLIBTABLE3="'+getField('CC_CODE')+'"' ;
if ExisteSQL (SQL) then begin LastError:=2; LastErrorMsg:=TexteMessage[LastError]; exit; end ;
End;

Procedure TOT_RTRPCLIBTABLE4.OnDeleteRecord;
var  SQL : string ;
Begin
// Test de l'existence du code dans les tables
SQL:='SELECT RPC_RPCLIBTABLE4 from PROSPECTCOMPL where RPC_RPCLIBTABLE4="'+getField('CC_CODE')+'"' ;
if ExisteSQL (SQL) then begin LastError:=2; LastErrorMsg:=TexteMessage[LastError]; exit; end ;
End;

Procedure TOT_RTRPCLIBTABLE5.OnDeleteRecord;
var  SQL : string ;
Begin
// Test de l'existence du code dans les tables
SQL:='SELECT RPC_RPCLIBTABLE5 from PROSPECTCOMPL where RPC_RPCLIBTABLE5="'+getField('CC_CODE')+'"' ;
if ExisteSQL (SQL) then begin LastError:=2; LastErrorMsg:=TexteMessage[LastError]; exit; end ;
End;

Procedure TOT_RTRPCLIBTABLE6.OnDeleteRecord;
var  SQL : string ;
Begin
// Test de l'existence du code dans les tables
SQL:='SELECT RPC_RPCLIBTABLE6 from PROSPECTCOMPL where RPC_RPCLIBTABLE6="'+getField('CC_CODE')+'"' ;
if ExisteSQL (SQL) then begin LastError:=2; LastErrorMsg:=TexteMessage[LastError]; exit; end ;
End;

Procedure TOT_RTRPCLIBTABLE7.OnDeleteRecord;
var  SQL : string ;
Begin
// Test de l'existence du code dans les tables
SQL:='SELECT RPC_RPCLIBTABLE7 from PROSPECTCOMPL where RPC_RPCLIBTABLE7="'+getField('CC_CODE')+'"' ;
if ExisteSQL (SQL) then begin LastError:=2; LastErrorMsg:=TexteMessage[LastError]; exit; end ;
End;

Procedure TOT_RTRPCLIBTABLE8.OnDeleteRecord;
var  SQL : string ;
Begin
// Test de l'existence du code dans les tables
SQL:='SELECT RPC_RPCLIBTABLE8 from PROSPECTCOMPL where RPC_RPCLIBTABLE8="'+getField('CC_CODE')+'"' ;
if ExisteSQL (SQL) then begin LastError:=2; LastErrorMsg:=TexteMessage[LastError]; exit; end ;
End;

Procedure TOT_RTRPCLIBTABLE9.OnDeleteRecord;
var  SQL : string ;
Begin
// Test de l'existence du code dans les tables
SQL:='SELECT RPC_RPCLIBTABLE9 from PROSPECTCOMPL where RPC_RPCLIBTABLE9="'+getField('CC_CODE')+'"' ;
if ExisteSQL (SQL) then begin LastError:=2; LastErrorMsg:=TexteMessage[LastError]; exit; end ;
End;

Procedure TOT_RTRPCLIBTABLEA.OnDeleteRecord;
var  SQL : string ;
Begin
// Test de l'existence du code dans les tables
SQL:='SELECT RPC_RPCLIBTABLEA from PROSPECTCOMPL where RPC_RPCLIBTABLEA="'+getField('CC_CODE')+'"' ;
if ExisteSQL (SQL) then begin LastError:=2; LastErrorMsg:=TexteMessage[LastError]; exit; end ;
End;

Procedure TOT_RTRPCLIBTABLEB.OnDeleteRecord;
var  SQL : string ;
Begin
// Test de l'existence du code dans les tables
SQL:='SELECT RPC_RPCLIBTABLEB from PROSPECTCOMPL where RPC_RPCLIBTABLEB="'+getField('CC_CODE')+'"' ;
if ExisteSQL (SQL) then begin LastError:=2; LastErrorMsg:=TexteMessage[LastError]; exit; end ;
End;

// ******* TABLE ACTIONS FOURNISSEURS *******
Procedure TOT_RFRPRLIBACTION1.OnDeleteRecord;
var  SQL : string ;
Begin
// Test de l'existence du code dans les tables
SQL:='SELECT RAC_TABLELIBREF1 from ACTIONS where RAC_TABLELIBREF1="'+getField('CC_CODE')+'"' ;
if ExisteSQL (SQL) then begin LastError:=2; LastErrorMsg:=TexteMessage[LastError]; exit; end ;
SQL:='SELECT RPA_TABLELIBREF1 from PARACTIONS where RPA_TABLELIBREF1="'+getField('CC_CODE')+'"' ;
if ExisteSQL (SQL) then begin LastError:=2; LastErrorMsg:=TexteMessage[LastError]; exit; end ;
SQL:='SELECT RAG_TABLELIBREF1 from ACTIONSGENERIQUES where RAG_TABLELIBREF1="'+getField('CC_CODE')+'"' ;
if ExisteSQL (SQL) then begin LastError:=2; LastErrorMsg:=TexteMessage[LastError]; exit; end ;
End;

Procedure TOT_RFRPRLIBACTION2.OnDeleteRecord;
var  SQL : string ;
Begin
// Test de l'existence du code dans les tables
SQL:='SELECT RAC_TABLELIBREF2 from ACTIONS where RAC_TABLELIBREF2="'+getField('CC_CODE')+'"' ;
if ExisteSQL (SQL) then begin LastError:=2; LastErrorMsg:=TexteMessage[LastError]; exit; end ;
SQL:='SELECT RPA_TABLELIBREF2 from PARACTIONS where RPA_TABLELIBREF2="'+getField('CC_CODE')+'"' ;
if ExisteSQL (SQL) then begin LastError:=2; LastErrorMsg:=TexteMessage[LastError]; exit; end ;
SQL:='SELECT RAG_TABLELIBREF2 from ACTIONSGENERIQUES where RAG_TABLELIBREF2="'+getField('CC_CODE')+'"' ;
if ExisteSQL (SQL) then begin LastError:=2; LastErrorMsg:=TexteMessage[LastError]; exit; end ;
End;

Procedure TOT_RFRPRLIBACTION3.OnDeleteRecord;
var  SQL : string ;
Begin
// Test de l'existence du code dans les tables
SQL:='SELECT RAC_TABLELIBREF3 from ACTIONS where RAC_TABLELIBREF3="'+getField('CC_CODE')+'"' ;
if ExisteSQL (SQL) then begin LastError:=2; LastErrorMsg:=TexteMessage[LastError]; exit; end ;
SQL:='SELECT RPA_TABLELIBREF3 from PARACTIONS where RPA_TABLELIBREF3="'+getField('CC_CODE')+'"' ;
if ExisteSQL (SQL) then begin LastError:=2; LastErrorMsg:=TexteMessage[LastError]; exit; end ;
SQL:='SELECT RAG_TABLELIBREF3 from ACTIONSGENERIQUES where RAG_TABLELIBREF3="'+getField('CC_CODE')+'"' ;
if ExisteSQL (SQL) then begin LastError:=2; LastErrorMsg:=TexteMessage[LastError]; exit; end ;
End;

// ******* TABLE CHAINAGES FOURNISSEURS *******
Procedure TOT_RFLIBCHAINAGE1.OnDeleteRecord;
var  SQL : string ;
Begin
// Test de l'existence du code dans les tables
SQL:='SELECT RCH_TABLELIBRECHF1 from ACTIONSCHAINEES where RCH_TABLELIBRECHF1="'+getField('CC_CODE')+'"' ;
if ExisteSQL (SQL) then begin LastError:=2; LastErrorMsg:=TexteMessage[LastError]; exit; end ;
SQL:='SELECT RPG_TABLELIBRECHF1 from PARCHAINAGES where RPG_TABLELIBRECHF1="'+getField('CC_CODE')+'"' ;
if ExisteSQL (SQL) then begin LastError:=2; LastErrorMsg:=TexteMessage[LastError]; exit; end ;
End;

Procedure TOT_RFLIBCHAINAGE2.OnDeleteRecord;
var  SQL : string ;
Begin
// Test de l'existence du code dans les tables
SQL:='SELECT RCH_TABLELIBRECHF2 from ACTIONSCHAINEES where RCH_TABLELIBRECHF2="'+getField('CC_CODE')+'"' ;
if ExisteSQL (SQL) then begin LastError:=2; LastErrorMsg:=TexteMessage[LastError]; exit; end ;
SQL:='SELECT RPG_TABLELIBRECHF2 from PARCHAINAGES where RPG_TABLELIBRECHF2="'+getField('CC_CODE')+'"' ;
if ExisteSQL (SQL) then begin LastError:=2; LastErrorMsg:=TexteMessage[LastError]; exit; end ;
End;

Procedure TOT_RFLIBCHAINAGE3.OnDeleteRecord;
var  SQL : string ;
Begin
// Test de l'existence du code dans les tables
SQL:='SELECT RCH_TABLELIBRECHF3 from ACTIONSCHAINEES where RCH_TABLELIBRECHF3="'+getField('CC_CODE')+'"' ;
if ExisteSQL (SQL) then begin LastError:=2; LastErrorMsg:=TexteMessage[LastError]; exit; end ;
SQL:='SELECT RPG_TABLELIBRECHF3 from PARCHAINAGES where RPG_TABLELIBRECHF3="'+getField('CC_CODE')+'"' ;
if ExisteSQL (SQL) then begin LastError:=2; LastErrorMsg:=TexteMessage[LastError]; exit; end ;
End;

// ******* TABLE OPERATIONS FOURNISSEURS *******
Procedure TOT_RFOBJETOPE.OnDeleteRecord;
var  SQL : string ;
Begin
// Test de l'existence du code dans les tables
SQL:='SELECT ROP_OBJETOPEF from OPERATIONS where ROP_OBJETOPEF="'+getField('CC_CODE')+'"' ;
if ExisteSQL (SQL) then begin LastError:=2; LastErrorMsg:=TexteMessage[LastError]; exit; end ;
End;

// ******* TABLE INFOS COMPLEMENTAIRES FOURNISSEURS *******
Procedure TOT_RT003LIBTABLE0.OnDeleteRecord;
var  SQL : string ;
Begin
// Test de l'existence du code dans les tables
SQL:='SELECT RD3_RD3LIBTABLE0 from RTINFOS003 where RD3_RD3LIBTABLE0="'+getField('CC_CODE')+'"' ;
if ExisteSQL (SQL) then begin LastError:=2; LastErrorMsg:=TexteMessage[LastError]; exit; end ;
End;

Procedure TOT_RT003LIBTABLE1.OnDeleteRecord;
var  SQL : string ;
Begin
// Test de l'existence du code dans les tables
SQL:='SELECT RD3_RD3LIBTABLE1 from RTINFOS003 where RD3_RD3LIBTABLE1="'+getField('CC_CODE')+'"' ;
if ExisteSQL (SQL) then begin LastError:=2; LastErrorMsg:=TexteMessage[LastError]; exit; end ;
End;

Procedure TOT_RT003LIBTABLE2.OnDeleteRecord;
var  SQL : string ;
Begin
// Test de l'existence du code dans les tables
SQL:='SELECT RD3_RD3LIBTABLE2 from RTINFOS003 where RD3_RD3LIBTABLE2="'+getField('CC_CODE')+'"' ;
if ExisteSQL (SQL) then begin LastError:=2; LastErrorMsg:=TexteMessage[LastError]; exit; end ;
End;

Procedure TOT_RT003LIBTABLE3.OnDeleteRecord;
var  SQL : string ;
Begin
// Test de l'existence du code dans les tables
SQL:='SELECT RD3_RD3LIBTABLE3 from RTINFOS003 where RD3_RD3LIBTABLE3="'+getField('CC_CODE')+'"' ;
if ExisteSQL (SQL) then begin LastError:=2; LastErrorMsg:=TexteMessage[LastError]; exit; end ;
End;

Procedure TOT_RT003LIBTABLE4.OnDeleteRecord;
var  SQL : string ;
Begin
// Test de l'existence du code dans les tables
SQL:='SELECT RD3_RD3LIBTABLE4 from RTINFOS003 where RD3_RD3LIBTABLE4="'+getField('CC_CODE')+'"' ;
if ExisteSQL (SQL) then begin LastError:=2; LastErrorMsg:=TexteMessage[LastError]; exit; end ;
End;

Procedure TOT_RT003LIBTABMUL0.OnDeleteRecord;
var  SQL : string ;
Begin
// Test de l'existence du code dans les tables
SQL:='SELECT RD3_RD3LIBMUL0 from RTINFOS003 where RD3_RD3LIBMUL0 like "'+getField('CC_CODE')+';%"' ;
if ExisteSQL (SQL) then begin LastError:=2; LastErrorMsg:=TexteMessage[LastError]; exit; end ;
SQL:='SELECT RD3_RD3LIBMUL0 from RTINFOS003 where RD3_RD3LIBMUL0 like "'+'%;'+getField('CC_CODE')+';%"' ;
if ExisteSQL (SQL) then begin LastError:=2; LastErrorMsg:=TexteMessage[LastError]; exit; end ;
End;

Procedure TOT_RT003LIBTABMUL1.OnDeleteRecord;
var  SQL : string ;
Begin
// Test de l'existence du code dans les tables
SQL:='SELECT RD3_RD3LIBMUL1 from RTINFOS003 where RD3_RD3LIBMUL1 like "'+getField('CC_CODE')+';%"' ;
if ExisteSQL (SQL) then begin LastError:=2; LastErrorMsg:=TexteMessage[LastError]; exit; end ;
SQL:='SELECT RD3_RD3LIBMUL1 from RTINFOS003 where RD3_RD3LIBMUL1 like "'+'%;'+getField('CC_CODE')+';%"' ;
if ExisteSQL (SQL) then begin LastError:=2; LastErrorMsg:=TexteMessage[LastError]; exit; end ;
End;

Procedure TOT_RT003LIBTABMUL2.OnDeleteRecord;
var  SQL : string ;
Begin
// Test de l'existence du code dans les tables
SQL:='SELECT RD3_RD3LIBMUL2 from RTINFOS003 where RD3_RD3LIBMUL2 like "'+getField('CC_CODE')+';%"' ;
if ExisteSQL (SQL) then begin LastError:=2; LastErrorMsg:=TexteMessage[LastError]; exit; end ;
SQL:='SELECT RD3_RD3LIBMUL2 from RTINFOS003 where RD3_RD3LIBMUL2 like "'+'%;'+getField('CC_CODE')+';%"' ;
if ExisteSQL (SQL) then begin LastError:=2; LastErrorMsg:=TexteMessage[LastError]; exit; end ;
End;

Procedure TOT_RT003LIBTABMUL3.OnDeleteRecord;
var  SQL : string ;
Begin
// Test de l'existence du code dans les tables
SQL:='SELECT RD3_RD3LIBMUL3 from RTINFOS003 where RD3_RD3LIBMUL3 like "'+getField('CC_CODE')+';%"' ;
if ExisteSQL (SQL) then begin LastError:=2; LastErrorMsg:=TexteMessage[LastError]; exit; end ;
SQL:='SELECT RD3_RD3LIBMUL3 from RTINFOS003 where RD3_RD3LIBMUL3 like "'+'%;'+getField('CC_CODE')+';%"' ;
if ExisteSQL (SQL) then begin LastError:=2; LastErrorMsg:=TexteMessage[LastError]; exit; end ;
End;

Procedure TOT_RT003LIBTABMUL4.OnDeleteRecord;
var  SQL : string ;
Begin
// Test de l'existence du code dans les tables
SQL:='SELECT RD3_RD3LIBMUL4 from RTINFOS003 where RD3_RD3LIBMUL4 like "'+getField('CC_CODE')+';%"' ;
if ExisteSQL (SQL) then begin LastError:=2; LastErrorMsg:=TexteMessage[LastError]; exit; end ;
SQL:='SELECT RD3_RD3LIBMUL4 from RTINFOS003 where RD3_RD3LIBMUL4 like "'+'%;'+getField('CC_CODE')+';%"' ;
if ExisteSQL (SQL) then begin LastError:=2; LastErrorMsg:=TexteMessage[LastError]; exit; end ;
End;

// ******* TABLE INFOS COMPLEMENTAIRES CONTACTS *******
Procedure TOT_RT006LIBTABLE0.OnDeleteRecord;
var  SQL : string ;
Begin
// Test de l'existence du code dans les tables
SQL:='SELECT RD6_RD6LIBTABLE0 from RTINFOS006 where RD6_RD6LIBTABLE0="'+getField('CC_CODE')+'"' ;
if ExisteSQL (SQL) then begin LastError:=2; LastErrorMsg:=TexteMessage[LastError]; exit; end ;
End;

Procedure TOT_RT006LIBTABLE1.OnDeleteRecord;
var  SQL : string ;
Begin
// Test de l'existence du code dans les tables
SQL:='SELECT RD6_RD6LIBTABLE1 from RTINFOS006 where RD6_RD6LIBTABLE1="'+getField('CC_CODE')+'"' ;
if ExisteSQL (SQL) then begin LastError:=2; LastErrorMsg:=TexteMessage[LastError]; exit; end ;
End;

Procedure TOT_RT006LIBTABLE2.OnDeleteRecord;
var  SQL : string ;
Begin
// Test de l'existence du code dans les tables
SQL:='SELECT RD6_RD6LIBTABLE2 from RTINFOS006 where RD6_RD6LIBTABLE2="'+getField('CC_CODE')+'"' ;
if ExisteSQL (SQL) then begin LastError:=2; LastErrorMsg:=TexteMessage[LastError]; exit; end ;
End;

Procedure TOT_RT006LIBTABLE3.OnDeleteRecord;
var  SQL : string ;
Begin
// Test de l'existence du code dans les tables
SQL:='SELECT RD6_RD6LIBTABLE3 from RTINFOS006 where RD6_RD6LIBTABLE3="'+getField('CC_CODE')+'"' ;
if ExisteSQL (SQL) then begin LastError:=2; LastErrorMsg:=TexteMessage[LastError]; exit; end ;
End;

Procedure TOT_RT006LIBTABLE4.OnDeleteRecord;
var  SQL : string ;
Begin
// Test de l'existence du code dans les tables
SQL:='SELECT RD6_RD6LIBTABLE4 from RTINFOS006 where RD6_RD6LIBTABLE4="'+getField('CC_CODE')+'"' ;
if ExisteSQL (SQL) then begin LastError:=2; LastErrorMsg:=TexteMessage[LastError]; exit; end ;
End;

Procedure TOT_RT006LIBTABMUL0.OnDeleteRecord;
var  SQL : string ;
Begin
// Test de l'existence du code dans les tables
SQL:='SELECT RD6_RD6LIBMUL0 from RTINFOS006 where RD6_RD6LIBMUL0 like "'+getField('CC_CODE')+';%"' ;
if ExisteSQL (SQL) then begin LastError:=2; LastErrorMsg:=TexteMessage[LastError]; exit; end ;
SQL:='SELECT RD6_RD6LIBMUL0 from RTINFOS006 where RD6_RD6LIBMUL0 like "'+'%;'+getField('CC_CODE')+';%"' ;
if ExisteSQL (SQL) then begin LastError:=2; LastErrorMsg:=TexteMessage[LastError]; exit; end ;
End;

Procedure TOT_RT006LIBTABMUL1.OnDeleteRecord;
var  SQL : string ;
Begin
// Test de l'existence du code dans les tables
SQL:='SELECT RD6_RD6LIBMUL1 from RTINFOS006 where RD6_RD6LIBMUL1 like "'+getField('CC_CODE')+';%"' ;
if ExisteSQL (SQL) then begin LastError:=2; LastErrorMsg:=TexteMessage[LastError]; exit; end ;
SQL:='SELECT RD6_RD6LIBMUL1 from RTINFOS006 where RD6_RD6LIBMUL1 like "'+'%;'+getField('CC_CODE')+';%"' ;
if ExisteSQL (SQL) then begin LastError:=2; LastErrorMsg:=TexteMessage[LastError]; exit; end ;
End;

Procedure TOT_RT006LIBTABMUL2.OnDeleteRecord;
var  SQL : string ;
Begin
// Test de l'existence du code dans les tables
SQL:='SELECT RD6_RD6LIBMUL2 from RTINFOS006 where RD6_RD6LIBMUL2 like "'+getField('CC_CODE')+';%"' ;
if ExisteSQL (SQL) then begin LastError:=2; LastErrorMsg:=TexteMessage[LastError]; exit; end ;
SQL:='SELECT RD6_RD6LIBMUL2 from RTINFOS006 where RD6_RD6LIBMUL2 like "'+'%;'+getField('CC_CODE')+';%"' ;
if ExisteSQL (SQL) then begin LastError:=2; LastErrorMsg:=TexteMessage[LastError]; exit; end ;
End;

Procedure TOT_RT006LIBTABMUL3.OnDeleteRecord;
var  SQL : string ;
Begin
// Test de l'existence du code dans les tables
SQL:='SELECT RD6_RD6LIBMUL3 from RTINFOS006 where RD6_RD6LIBMUL3 like "'+getField('CC_CODE')+';%"' ;
if ExisteSQL (SQL) then begin LastError:=2; LastErrorMsg:=TexteMessage[LastError]; exit; end ;
SQL:='SELECT RD6_RD6LIBMUL3 from RTINFOS006 where RD6_RD6LIBMUL3 like "'+'%;'+getField('CC_CODE')+';%"' ;
if ExisteSQL (SQL) then begin LastError:=2; LastErrorMsg:=TexteMessage[LastError]; exit; end ;
End;

Procedure TOT_RT006LIBTABMUL4.OnDeleteRecord;
var  SQL : string ;
Begin
// Test de l'existence du code dans les tables
SQL:='SELECT RD6_RD6LIBMUL4 from RTINFOS006 where RD6_RD6LIBMUL4 like "'+getField('CC_CODE')+';%"' ;
if ExisteSQL (SQL) then begin LastError:=2; LastErrorMsg:=TexteMessage[LastError]; exit; end ;
SQL:='SELECT RD6_RD6LIBMUL4 from RTINFOS006 where RD6_RD6LIBMUL4 like "'+'%;'+getField('CC_CODE')+';%"' ;
if ExisteSQL (SQL) then begin LastError:=2; LastErrorMsg:=TexteMessage[LastError]; exit; end ;
End;

{ TOT_RTCORRESPIMPORT }
{$IFDEF SAV}
Procedure TOT_WLIBREWIV1.OnDeleteRecord;
var  SQL : string ;
Begin
SQL:='SELECT WIV_LIBREWIV1 from WINTERVENTION where WIV_LIBREWIV1="'+getField('YX_CODE')+'"' ;
if ExisteSQL (SQL) then begin LastError:=2; LastErrorMsg:=TexteMessage[LastError]; exit; end ;
End;

Procedure TOT_WLIBREWIV2.OnDeleteRecord;
var  SQL : string ;
Begin
SQL:='SELECT WIV_LIBREWIV2 from WINTERVENTION where WIV_LIBREWIV2="'+getField('YX_CODE')+'"' ;
if ExisteSQL (SQL) then begin LastError:=2; LastErrorMsg:=TexteMessage[LastError]; exit; end ;
End;

Procedure TOT_WLIBREWIV3.OnDeleteRecord;
var  SQL : string ;
Begin
SQL:='SELECT WIV_LIBREWIV3 from WINTERVENTION where WIV_LIBREWIV3="'+getField('YX_CODE')+'"' ;
if ExisteSQL (SQL) then begin LastError:=2; LastErrorMsg:=TexteMessage[LastError]; exit; end ;
End;

Procedure TOT_WLIBREWIV4.OnDeleteRecord;
var  SQL : string ;
Begin
SQL:='SELECT WIV_LIBREWIV4 from WINTERVENTION where WIV_LIBREWIV4="'+getField('YX_CODE')+'"' ;
if ExisteSQL (SQL) then begin LastError:=2; LastErrorMsg:=TexteMessage[LastError]; exit; end ;
End;

Procedure TOT_WLIBREWIV5.OnDeleteRecord;
var  SQL : string ;
Begin
SQL:='SELECT WIV_LIBREWIV5 from WINTERVENTION where WIV_LIBREWIV5="'+getField('YX_CODE')+'"' ;
if ExisteSQL (SQL) then begin LastError:=2; LastErrorMsg:=TexteMessage[LastError]; exit; end ;
End;

Procedure TOT_WLIBREWIV6.OnDeleteRecord;
var  SQL : string ;
Begin
SQL:='SELECT WIV_LIBREWIV6 from WINTERVENTION where WIV_LIBREWIV6="'+getField('YX_CODE')+'"' ;
if ExisteSQL (SQL) then begin LastError:=2; LastErrorMsg:=TexteMessage[LastError]; exit; end ;
End;

Procedure TOT_WLIBREWIV7.OnDeleteRecord;
var  SQL : string ;
Begin
SQL:='SELECT WIV_LIBREWIV7 from WINTERVENTION where WIV_LIBREWIV7="'+getField('YX_CODE')+'"' ;
if ExisteSQL (SQL) then begin LastError:=2; LastErrorMsg:=TexteMessage[LastError]; exit; end ;
End;

Procedure TOT_WLIBREWIV8.OnDeleteRecord;
var  SQL : string ;
Begin
SQL:='SELECT WIV_LIBREWIV8 from WINTERVENTION where WIV_LIBREWIV8="'+getField('YX_CODE')+'"' ;
if ExisteSQL (SQL) then begin LastError:=2; LastErrorMsg:=TexteMessage[LastError]; exit; end ;
End;

Procedure TOT_WLIBREWIV9.OnDeleteRecord;
var  SQL : string ;
Begin
SQL:='SELECT WIV_LIBREWIV9 from WINTERVENTION where WIV_LIBREWIV9="'+getField('YX_CODE')+'"' ;
if ExisteSQL (SQL) then begin LastError:=2; LastErrorMsg:=TexteMessage[LastError]; exit; end ;
End;

Procedure TOT_WLIBREWIVA.OnDeleteRecord;
var  SQL : string ;
Begin
SQL:='SELECT WIV_LIBREWIVA from WINTERVENTION where WIV_LIBREWIVA="'+getField('YX_CODE')+'"' ;
if ExisteSQL (SQL) then begin LastError:=2; LastErrorMsg:=TexteMessage[LastError]; exit; end ;
End;

Procedure TOT_RT007LIBTABLE1.OnDeleteRecord;
var  SQL : string ;
Begin
SQL:='SELECT RD7_RD7LIBTABLE1 from RTINFOS007 where RD7_RD7LIBTABLE1="'+getField('CC_CODE')+'"' ;
if ExisteSQL (SQL) then begin LastError:=2; LastErrorMsg:=TexteMessage[LastError]; exit; end ;
End;

Procedure TOT_RT007LIBTABLE2.OnDeleteRecord;
var  SQL : string ;
Begin
SQL:='SELECT RD7_RD7LIBTABLE2 from RTINFOS007 where RD7_RD7LIBTABLE2="'+getField('CC_CODE')+'"' ;
if ExisteSQL (SQL) then begin LastError:=2; LastErrorMsg:=TexteMessage[LastError]; exit; end ;
End;

Procedure TOT_RT007LIBTABLE3.OnDeleteRecord;
var  SQL : string ;
Begin
SQL:='SELECT RD7_RD7LIBTABLE3 from RTINFOS007 where RD7_RD7LIBTABLE3="'+getField('CC_CODE')+'"' ;
if ExisteSQL (SQL) then begin LastError:=2; LastErrorMsg:=TexteMessage[LastError]; exit; end ;
End;

Procedure TOT_RT007LIBTABLE4.OnDeleteRecord;
var  SQL : string ;
Begin
SQL:='SELECT RD7_RD7LIBTABLE4 from RTINFOS007 where RD7_RD7LIBTABLE4="'+getField('CC_CODE')+'"' ;
if ExisteSQL (SQL) then begin LastError:=2; LastErrorMsg:=TexteMessage[LastError]; exit; end ;
End;

Procedure TOT_RT007LIBTABLE0.OnDeleteRecord;
var  SQL : string ;
Begin
SQL:='SELECT RD7_RD7LIBTABLE0 from RTINFOS007 where RD7_RD7LIBTABLE0="'+getField('CC_CODE')+'"' ;
if ExisteSQL (SQL) then begin LastError:=2; LastErrorMsg:=TexteMessage[LastError]; exit; end ;
End;
Procedure TOT_RT007LIBTABMUL1.OnDeleteRecord;
var  SQL : string ;
Begin
SQL:='SELECT RD7_RD7LIBMUL1 from RTINFOS007 where RD7_RD7LIBMUL1="'+getField('CC_CODE')+'"' ;
if ExisteSQL (SQL) then begin LastError:=2; LastErrorMsg:=TexteMessage[LastError]; exit; end ;
End;

Procedure TOT_RT007LIBTABMUL2.OnDeleteRecord;
var  SQL : string ;
Begin
SQL:='SELECT RD7_RD7LIBMUL2 from RTINFOS007 where RD7_RD7LIBMUL2="'+getField('CC_CODE')+'"' ;
if ExisteSQL (SQL) then begin LastError:=2; LastErrorMsg:=TexteMessage[LastError]; exit; end ;
End;

Procedure TOT_RT007LIBTABMUL3.OnDeleteRecord;
var  SQL : string ;
Begin
SQL:='SELECT RD7_RD7LIBMUL3 from RTINFOS007 where RD7_RD7LIBMUL3="'+getField('CC_CODE')+'"' ;
if ExisteSQL (SQL) then begin LastError:=2; LastErrorMsg:=TexteMessage[LastError]; exit; end ;
End;

Procedure TOT_RT007LIBTABMUL4.OnDeleteRecord;
var  SQL : string ;
Begin
SQL:='SELECT RD7_RD7LIBMUL4 from RTINFOS007 where RD7_RD7LIBMUL4="'+getField('CC_CODE')+'"' ;
if ExisteSQL (SQL) then begin LastError:=2; LastErrorMsg:=TexteMessage[LastError]; exit; end ;
End;

Procedure TOT_RT007LIBTABMUL0.OnDeleteRecord;
var  SQL : string ;
Begin
SQL:='SELECT RD7_RD7LIBMUL0 from RTINFOS007 where RD7_RD7LIBMUL0="'+getField('CC_CODE')+'"' ;
if ExisteSQL (SQL) then begin LastError:=2; LastErrorMsg:=TexteMessage[LastError]; exit; end ;
End;

{$ENDIF SAV}
Procedure TOT_RT00QLIBTABLE0.OnDeleteRecord;
var  SQL : string ;
Begin
SQL:='SELECT RDQ_RDQLIBTABLE0 from RTINFOS00Q where RDQ_RDQLIBTABLE0="'+getField('CC_CODE')+'"' ;
if ExisteSQL (SQL) then begin LastError:=2; LastErrorMsg:=TexteMessage[LastError]; exit; end ;
End;

Procedure TOT_RT00QLIBTABLE1.OnDeleteRecord;
var  SQL : string ;
Begin
SQL:='SELECT RDQ_RDQLIBTABLE1 from RTINFOS00Q where RDQ_RDQLIBTABLE1="'+getField('CC_CODE')+'"' ;
if ExisteSQL (SQL) then begin LastError:=2; LastErrorMsg:=TexteMessage[LastError]; exit; end ;
End;

Procedure TOT_RT00QLIBTABLE2.OnDeleteRecord;
var  SQL : string ;
Begin
SQL:='SELECT RDQ_RDQLIBTABLE2 from RTINFOS00Q where RDQ_RDQLIBTABLE2="'+getField('CC_CODE')+'"' ;
if ExisteSQL (SQL) then begin LastError:=2; LastErrorMsg:=TexteMessage[LastError]; exit; end ;
End;

Procedure TOT_RT00QLIBTABLE3.OnDeleteRecord;
var  SQL : string ;
Begin
SQL:='SELECT RDQ_RDQLIBTABLE3 from RTINFOS00Q where RDQ_RDQLIBTABLE3="'+getField('CC_CODE')+'"' ;
if ExisteSQL (SQL) then begin LastError:=2; LastErrorMsg:=TexteMessage[LastError]; exit; end ;
End;

Procedure TOT_RT00QLIBTABLE4.OnDeleteRecord;
var  SQL : string ;
Begin
SQL:='SELECT RDQ_RDQLIBTABLE4 from RTINFOS00Q where RDQ_RDQLIBTABLE4="'+getField('CC_CODE')+'"' ;
if ExisteSQL (SQL) then begin LastError:=2; LastErrorMsg:=TexteMessage[LastError]; exit; end ;
End;

Procedure TOT_RT00QLIBTABMUL0.OnDeleteRecord;
var  SQL : string ;
Begin
SQL:='SELECT RDQ_RDQLIBMUL0 from RTINFOS00Q where RDQ_RDQLIBMUL0 like "'+getField('CC_CODE')+';%"' ;
if ExisteSQL (SQL) then begin LastError:=2; LastErrorMsg:=TexteMessage[LastError]; exit; end ;
SQL:='SELECT RDQ_RDQLIBMUL0 from RTINFOS00Q where RDQ_RDQLIBMUL0 like "'+'%;'+getField('CC_CODE')+';%"' ;
if ExisteSQL (SQL) then begin LastError:=2; LastErrorMsg:=TexteMessage[LastError]; exit; end ;
End;

Procedure TOT_RT00QLIBTABMUL1.OnDeleteRecord;
var  SQL : string ;
Begin
SQL:='SELECT RDQ_RDQLIBMUL1 from RTINFOS00Q where RDQ_RDQLIBMUL1 like "'+getField('CC_CODE')+';%"' ;
if ExisteSQL (SQL) then begin LastError:=2; LastErrorMsg:=TexteMessage[LastError]; exit; end ;
SQL:='SELECT RDQ_RDQLIBMUL1 from RTINFOS00Q where RDQ_RDQLIBMUL1 like "'+'%;'+getField('CC_CODE')+';%"' ;
if ExisteSQL (SQL) then begin LastError:=2; LastErrorMsg:=TexteMessage[LastError]; exit; end ;
End;

Procedure TOT_RT00QLIBTABMUL2.OnDeleteRecord;
var  SQL : string ;
Begin
SQL:='SELECT RDQ_RDQLIBMUL2 from RTINFOS00Q where RDQ_RDQLIBMUL2 like "'+getField('CC_CODE')+';%"' ;
if ExisteSQL (SQL) then begin LastError:=2; LastErrorMsg:=TexteMessage[LastError]; exit; end ;
SQL:='SELECT RDQ_RDQLIBMUL2 from RTINFOS00Q where RDQ_RDQLIBMUL2 like "'+'%;'+getField('CC_CODE')+';%"' ;
if ExisteSQL (SQL) then begin LastError:=2; LastErrorMsg:=TexteMessage[LastError]; exit; end ;
End;

Procedure TOT_RT00QLIBTABMUL3.OnDeleteRecord;
var  SQL : string ;
Begin
SQL:='SELECT RDQ_RDQLIBMUL3 from RTINFOS00Q where RDQ_RDQLIBMUL3 like "'+getField('CC_CODE')+';%"' ;
if ExisteSQL (SQL) then begin LastError:=2; LastErrorMsg:=TexteMessage[LastError]; exit; end ;
SQL:='SELECT RDQ_RDQLIBMUL3 from RTINFOS00Q where RDQ_RDQLIBMUL3 like "'+'%;'+getField('CC_CODE')+';%"' ;
if ExisteSQL (SQL) then begin LastError:=2; LastErrorMsg:=TexteMessage[LastError]; exit; end ;
End;

Procedure TOT_RT00QLIBTABMUL4.OnDeleteRecord;
var  SQL : string ;
Begin
SQL:='SELECT RDQ_RDQLIBMUL4 from RTINFOS00Q where RDQ_RDQLIBMUL4 like "'+getField('CC_CODE')+';%"' ;
if ExisteSQL (SQL) then begin LastError:=2; LastErrorMsg:=TexteMessage[LastError]; exit; end ;
SQL:='SELECT RDQ_RDQLIBMUL4 from RTINFOS00Q where RDQ_RDQLIBMUL4 like "'+'%;'+getField('CC_CODE')+';%"' ;
if ExisteSQL (SQL) then begin LastError:=2; LastErrorMsg:=TexteMessage[LastError]; exit; end ;
End;

Procedure TOT_RT00VLIBTABLE0.OnDeleteRecord;
var  SQL : string ;
Begin
SQL:='SELECT RDV_RDVLIBTABLE0 from RTINFOS00V where RDV_RDVLIBTABLE0="'+getField('CC_CODE')+'"' ;
if ExisteSQL (SQL) then begin LastError:=2; LastErrorMsg:=TexteMessage[LastError]; exit; end ;
End;

Procedure TOT_RT00VLIBTABLE1.OnDeleteRecord;
var  SQL : string ;
Begin
SQL:='SELECT RDV_RDVLIBTABLE1 from RTINFOS00V where RDV_RDVLIBTABLE1="'+getField('CC_CODE')+'"' ;
if ExisteSQL (SQL) then begin LastError:=2; LastErrorMsg:=TexteMessage[LastError]; exit; end ;
End;

Procedure TOT_RT00VLIBTABLE2.OnDeleteRecord;
var  SQL : string ;
Begin
SQL:='SELECT RDV_RDVLIBTABLE2 from RTINFOS00V where RDV_RDVLIBTABLE2="'+getField('CC_CODE')+'"' ;
if ExisteSQL (SQL) then begin LastError:=2; LastErrorMsg:=TexteMessage[LastError]; exit; end ;
End;

Procedure TOT_RT00VLIBTABLE3.OnDeleteRecord;
var  SQL : string ;
Begin
SQL:='SELECT RDV_RDVLIBTABLE3 from RTINFOS00V where RDV_RDVLIBTABLE3="'+getField('CC_CODE')+'"' ;
if ExisteSQL (SQL) then begin LastError:=2; LastErrorMsg:=TexteMessage[LastError]; exit; end ;
End;

Procedure TOT_RT00VLIBTABLE4.OnDeleteRecord;
var  SQL : string ;
Begin
SQL:='SELECT RDV_RDVLIBTABLE4 from RTINFOS00V where RDV_RDVLIBTABLE4="'+getField('CC_CODE')+'"' ;
if ExisteSQL (SQL) then begin LastError:=2; LastErrorMsg:=TexteMessage[LastError]; exit; end ;
End;

Procedure TOT_RT00VLIBTABMUL0.OnDeleteRecord;
var  SQL : string ;
Begin
SQL:='SELECT RDV_RDVLIBMUL0 from RTINFOS00V where RDV_RDVLIBMUL0 like "'+getField('CC_CODE')+';%"' ;
if ExisteSQL (SQL) then begin LastError:=2; LastErrorMsg:=TexteMessage[LastError]; exit; end ;
SQL:='SELECT RDV_RDVLIBMUL0 from RTINFOS00V where RDV_RDVLIBMUL0 like "'+'%;'+getField('CC_CODE')+';%"' ;
if ExisteSQL (SQL) then begin LastError:=2; LastErrorMsg:=TexteMessage[LastError]; exit; end ;
End;

Procedure TOT_RT00VLIBTABMUL1.OnDeleteRecord;
var  SQL : string ;
Begin
SQL:='SELECT RDV_RDVLIBMUL1 from RTINFOS00V where RDV_RDVLIBMUL1 like "'+getField('CC_CODE')+';%"' ;
if ExisteSQL (SQL) then begin LastError:=2; LastErrorMsg:=TexteMessage[LastError]; exit; end ;
SQL:='SELECT RDV_RDVLIBMUL1 from RTINFOS00V where RDV_RDVLIBMUL1 like "'+'%;'+getField('CC_CODE')+';%"' ;
if ExisteSQL (SQL) then begin LastError:=2; LastErrorMsg:=TexteMessage[LastError]; exit; end ;
End;

Procedure TOT_RT00VLIBTABMUL2.OnDeleteRecord;
var  SQL : string ;
Begin
SQL:='SELECT RDV_RDVLIBMUL2 from RTINFOS00V where RDV_RDVLIBMUL2 like "'+getField('CC_CODE')+';%"' ;
if ExisteSQL (SQL) then begin LastError:=2; LastErrorMsg:=TexteMessage[LastError]; exit; end ;
SQL:='SELECT RDV_RDVLIBMUL2 from RTINFOS00V where RDV_RDVLIBMUL2 like "'+'%;'+getField('CC_CODE')+';%"' ;
if ExisteSQL (SQL) then begin LastError:=2; LastErrorMsg:=TexteMessage[LastError]; exit; end ;
End;

Procedure TOT_RT00VLIBTABMUL3.OnDeleteRecord;
var  SQL : string ;
Begin
SQL:='SELECT RDV_RDVLIBMUL3 from RTINFOS00V where RDV_RDVLIBMUL3 like "'+getField('CC_CODE')+';%"' ;
if ExisteSQL (SQL) then begin LastError:=2; LastErrorMsg:=TexteMessage[LastError]; exit; end ;
SQL:='SELECT RDV_RDVLIBMUL3 from RTINFOS00V where RDV_RDVLIBMUL3 like "'+'%;'+getField('CC_CODE')+';%"' ;
if ExisteSQL (SQL) then begin LastError:=2; LastErrorMsg:=TexteMessage[LastError]; exit; end ;
End;

Procedure TOT_RT00VLIBTABMUL4.OnDeleteRecord;
var  SQL : string ;
Begin
SQL:='SELECT RDV_RDVLIBMUL4 from RTINFOS00V where RDV_RDVLIBMUL4 like "'+getField('CC_CODE')+';%"' ;
if ExisteSQL (SQL) then begin LastError:=2; LastErrorMsg:=TexteMessage[LastError]; exit; end ;
SQL:='SELECT RDV_RDVLIBMUL4 from RTINFOS00V where RDV_RDVLIBMUL4 like "'+'%;'+getField('CC_CODE')+';%"' ;
if ExisteSQL (SQL) then begin LastError:=2; LastErrorMsg:=TexteMessage[LastError]; exit; end ;
End;


procedure TOT_RTCORRESPIMPORT.FlisteOnElipsisClick ( sender : Tobject) ;
var
  LaTable               : String;
  Prefixe : String;
  Tablette : String;
  Numtab : integer;
  VPDE : TDECombo;
  MenuTag : integer;
  LeChamp               : String;
  TobParSpctCor         : Tob;

begin
{$IFDEF EAGLCLIENT}
  LeChamp               := THGrid(TFTablette(Ecran).Fliste).Cells[3, THGrid(TFTablette(Ecran).Fliste).Row];
  TobParSpctCor         := nil;
  LaTable               := RenvoiTabletteCor(TobParSpctCor, LeChamp);
  if LaTable = '' then
    LaTable := ChampToTT(LeChamp);
{$ELSE}
//  LaTable := ChampToTT(TFTablette(Ecran).Fliste.DataSource.DataSet.Fields[4].value);
  LeChamp               := Getfield('CC_LIBRE');
  LaTable               := RenvoiTabletteCor(TobParSpctCor, LeChamp);
  if LaTable = '' then
    LaTable := ChampToTT(LeChamp);
{$ENDIF}

  Numtab := TTToNum(LaTable);
  if Numtab <= 0 then
    exit;

  VPDE := V_PGI.DECombos[Numtab];
  Prefixe := TTToPrefixe(LaTable);
  if prefixe = 'CC' then
    Tablette := 'CHOIXCOD'
  else if Prefixe = 'PY' then
    Tablette := 'PAYS'
  else if Prefixe = 'CO' then
    Tablette := 'COMMUN'
  else if Prefixe = 'YX' then
    Tablette := 'CHOIXEXT'
  else if Prefixe = 'RG' then
    Tablette := 'REGION';

//  if (VPDE.TagModif = 0) or (VPDE.TagModif = 999) then
    MenuTag := -1;
//  else
//    MenuTag := VPDE.TagModif;

  FreeAndNil(TobParSpctCor);
{$IFDEF EAGLCLIENT}
  LookupList(THGrid(Sender).InplaceEditor, VPDE.Libelle, Tablette, VPDE.code, VPDE.ChampLib, VPDE.where, VPDE.code, False, MenuTag);
{$ELSE}
  THDBGrid(Sender).DataSource.Edit;
  if LookupList(THDBGrid(Sender).InplaceEditor, VPDE.libelle, Tablette, VPDE.code, VPDE.ChampLib, VPDE.where, VPDE.code, False, MenuTag) then
    THDBGrid(Sender).SelectedField.AsString := THDBGrid(Sender).InplaceEditor.Text;
{$ENDIF}

end;


procedure TOT_RTCORRESPIMPORT.FListeOnsorted(Sender: Tobject);
var
  LesColonnes : String;

begin
  LesColonnes := 'CC_CODE;CC_LIBELLE;CC_ABREGE;CC_LIBRE;CC_TYPE;LIBCHAMP';

{$IFDEF EAGLCLIENT}
  Tob(Self.DS).GetGridDetail(TFTablette(Ecran).Fliste, THGrid(TFTablette(Ecran).Fliste).RowCount -1, 'CHOIXCOD', LesColonnes, True);
{$ENDIF}

end;

procedure TOT_RTCORRESPIMPORT.OnArgument ( S : String) ;
var
  LesColonnes : String;
{$IFDEF EAGLCLIENT}
  i : integer;
{  LaTable : String;
  Numtab : integer;
  VPDE : TDECombo;
}
  LeLib : String;
{$ENDIF EAGLCLIENT}

begin
  inherited;
  LesColonnes := 'CC_CODE;CC_LIBELLE;CC_ABREGE;CC_LIBRE;CC_TYPE;LIBCHAMP';

{$IFDEF EAGLCLIENT}
  TFTablette(Ecran).Fliste.Colcount := 6;

  TFTablette(Ecran).fliste.FColTypes[3] := 'S';
  TFTablette(Ecran).fliste.Titres.Add ('Libres');
  TFTablette(Ecran).fliste.Titres.Add ('Champs');
  TFTablette(Ecran).fliste.Cells[0,0]  := 'Code';
  TFTablette(Ecran).fliste.Cells[1,0]  := 'Libellé fichier import';
  TFTablette(Ecran).fliste.Cells[2,0]  := 'Code associé';
  TFTablette(Ecran).fliste.Cells[3,0]  := 'Champ associé';
  TFTablette(Ecran).fliste.Cells[4,0]  := 'Type';
  TFTablette(Ecran).fliste.Cells[5,0]  := 'Libellé champ associé';

  if TFTablette(Ecran).FListe.RowCount > 2 then
  begin
    for i := 0 to TFTablette(Ecran).FListe.RowCount -2 do
    begin
  {    LaTable := ChampToTT(Self.DS.Detail[i].Getvalue('CC_LIBRE'));
      Numtab := TTToNum(LaTable);
      if Numtab <= 0 then
        exit;
      VPDE := V_PGI.DECombos[Numtab];
      Self.DS.Detail[i].AddChampSupValeur('LIBCHAMP', VPDE.libelle);
      }
      LeLib := ChampToLibelle(Self.DS.Detail[i].Getvalue('CC_LIBRE'));
      Self.DS.Detail[i].AddChampSupValeur('LIBCHAMP', LeLib);

    end;
  end;
  Tob(Self.ds).PutGridDetail(TFTablette(Ecran).Fliste,False,False,LesColonnes);
  TFTablette(Ecran).FListe.ColWidths[0] := 50;  //code
  TFTablette(Ecran).FListe.ColWidths[1] := 200;  //Libellé import
  TFTablette(Ecran).FListe.ColWidths[2] := 70;  // Code concordance
  TFTablette(Ecran).FListe.ColWidths[3] := -1;  // Champ associé
  TFTablette(Ecran).fliste.ColWidths[4] := -1;  //Type
  TFTablette(Ecran).fliste.ColWidths[5] := 150; // libellé champ associé

  TFTablette(Ecran).HMTrad.ResizeGridColumns(TFTablette(Ecran).Fliste);

  THGrid(TFTablette(Ecran).Fliste).ColEditables[0] := False;
  THGrid(TFTablette(Ecran).Fliste).ColEditables[1] := False;
  THGrid(TFTablette(Ecran).Fliste).ColEditables[3] := False;
  THGrid(TFTablette(Ecran).Fliste).ColEditables[4] := False;
  THGrid(TFTablette(Ecran).Fliste).ColEditables[5] := False;
  THGrid(Tftablette(ecran).fliste).ElipsisButton := True;
  THGrid(Tftablette(ecran).fliste).ElipsisHint := 'Recherche dans la tablette associée';
  THGrid(Tftablette(ecran).fliste).OnElipsisClick := FlisteOnElipsisClick;
  THGrid(TFTablette(Ecran).Fliste).OnSorted := FListeOnsorted;
  TFTablette(Ecran).FListe.SortEnabled := True;
  TFTablette(Ecran).FListe.SortGrid(3, True);



{$ELSE}

  TFTablette(Ecran).FListe.Columns.Add;
  TFTablette(Ecran).FListe.Columns[3].FieldName := 'CC_LIBRE';
  TFTablette(Ecran).FListe.Columns[3].DisplayName := 'CC_LIBRE';

  TFTablette(Ecran).FListe.Columns[0].Title.Caption := 'Code';
  TFTablette(Ecran).FListe.Columns[1].Title.Caption := 'Libellé fichier import';
  TFTablette(Ecran).FListe.Columns[2].Title.Caption := 'Code associé';
  TFTablette(Ecran).FListe.Columns[3].Title.Caption := 'Champ associé';
  TFTablette(Ecran).FListe.Columns[1].Width := 250;
  TFTablette(Ecran).FListe.Columns[2].Width := 70;
  TFTablette(Ecran).FListe.Columns[3].Width := 150;
  TFTablette(Ecran).FListe.Columns[1].ReadOnly := True;
  TFTablette(Ecran).FListe.Columns[3].ReadOnly := True;

  THDBGrid(TFTablette(Ecran).Fliste).SortEnabled := True;
  THDBGrid(TFTablette(Ecran).FListe).SortingColumns := 'CC_LIBRE';
  THDBGrid(TFTablette(Ecran).FListe).Columns[2].ButtonStyle := cbsEllipsis;
  THDBGrid(TFTablette(Ecran).FListe).OnEditButtonClick := FlisteOnElipsisClick;
//  TFTablette(Ecran).TChoixCod.IndexFieldNames := 'CC_LIBRE';
//  TFTablette(Ecran).TChoixCod.IndexName := 'CC_LIBRE';

{$ENDIF}

  TFTablette(Ecran).BDelete.Enabled := True;
  TFTablette(Ecran).BDelete.Visible := True;

  self.OkComplement := true;
  TFTablette(Ecran).bComplement.visible := true;
  if TFTablette(Ecran).bComplement <> nil then
    TFTablette(Ecran).bComplement.onClick := bComplementClick;

end;

procedure TOT_RTCORRESPIMPORT.OnDeleteRecord;
var
  LesColonnes : String;

begin
  LesColonnes := 'CC_CODE;CC_LIBELLE;CC_ABREGE;CC_LIBRE;CC_TYPE;LIBCHAMP';

{$IFDEF EAGLCLIENT}
  Tob(Self.DS).GetGridDetail(TFTablette(Ecran).Fliste, THGrid(TFTablette(Ecran).Fliste).RowCount -1, 'CHOIXCOD', LesColonnes, True);
{$ENDIF}
  inherited;

end;


procedure TOT_RTCORRESPIMPORT.bComplementClick(Sender: TObject);
var
//{$IFNDEF EAGLCLIENT}
  LeChamp               : String;
//{$ENDIF}
  LaTable : String;
  TobParSuspectCor      : Tob;

begin
  LaTable := '';
  {$IFDEF EAGLCLIENT}
  LeChamp               := THGrid(TFTablette(Ecran).Fliste).Cells[3, THGrid(TFTablette(Ecran).Fliste).Row];
  // on recherche s'il y a une correspondance avec une tablette tiers
  LaTable               := RenvoiTabletteCor(TobParSuspectCor, LeChamp);
    // sinon la tablette lié au champ
  if LaTable = '' then
    LaTable             := ChampToTT(LeChamp);
  {$ELSE}
  THDBGrid(TFTablette(Ecran).Fliste).SortEnabled := True;
  //  LaTable := ChampToTT(TFTablette(Ecran).Fliste.DataSource.DataSet.Fields[4].value);
  LeChamp := Getfield('CC_LIBRE');
  LaTable               := RenvoiTabletteCor(TobParSuspectCor, LeChamp);
    // sinon la tablette lié au champ
  if LaTable = '' then
    LaTable := ChampToTT(LeChamp);

  {$ENDIF}
  if (LaTable <> '') then
  begin
    if LaTable = 'TTPAYS' then
      AGLLanceFiche('YY', 'YYPAYS', '', '', '')
    else if LaTable = 'GCREGION' then
      AGLLanceFiche('GC', 'GCREGION_FSL', '', '', '')
    else
      ParamTable(LaTable , taCreat, -1, nil);
  end;
  FreeAndNil(TobParSuspectCor);
end;

Procedure TOT_RBCATEGORIE.OnDeleteRecord;
var  SQL : string ;
Begin
SQL:='SELECT BCO_CATEGORIEBC FROM BASECONNAISSANCE WHERE BCO_CATEGORIEBC="'+getField('CC_CODE')+'"' ;
if ExisteSQL (SQL) then begin LastError:=2; LastErrorMsg:=TexteMessage[LastError]; exit; end ;
End;

Procedure TOT_RBMULTICHOIXBC1.OnDeleteRecord;
var  SQL : string ;
Begin
SQL:='SELECT BCO_CHARLIBRE1 FROM BASECONNAISSANCE WHERE BCO_CHARLIBRE1 LIKE "%'+getField('CC_CODE')+';%"' ;
if ExisteSQL (SQL) then begin LastError:=2; LastErrorMsg:=TexteMessage[LastError]; exit; end ;
End;

Procedure TOT_RBTYPEINFORMATION.OnDeleteRecord;
var  SQL : string ;
Begin
SQL:='SELECT BCO_TYPEINFORMATION FROM BASECONNAISSANCE WHERE BCO_TYPEINFORMATION="'+getField('CC_CODE')+'"' ;
if ExisteSQL (SQL) then begin LastError:=2; LastErrorMsg:=TexteMessage[LastError]; exit; end ;
End;

Initialization
registerclasses([TOT_RTMOTIFS,TOT_RTTYPEPERSPECTIVE]) ;
registerclasses([TOT_RTRPRLIBPERSPECTIVE1,TOT_RTRPRLIBPERSPECTIVE2,TOT_RTRPRLIBPERSPECTIVE3]) ;
registerclasses([TOT_RTIMPORTANCEACT]) ;
registerclasses([TOT_RTRPRLIBACTION1,TOT_RTRPRLIBACTION2,TOT_RTRPRLIBACTION3]) ;
registerclasses([TOT_RT001LIBTABLE0,TOT_RT001LIBTABLE1,TOT_RT001LIBTABLE2]) ;
registerclasses([TOT_RT001LIBTABLE3,TOT_RT001LIBTABLE4]) ;
registerclasses([TOT_RT001LIBTABMUL0,TOT_RT001LIBTABMUL1,TOT_RT001LIBTABMUL2]) ;
registerclasses([TOT_RT001LIBTABMUL3,TOT_RT001LIBTABMUL4]) ;
registerclasses([TOT_RTLIBCHAINAGE1,TOT_RTLIBCHAINAGE2,TOT_RTLIBCHAINAGE3]) ;
registerclasses([TOT_RTRSCLIBTABLE0,TOT_RTRSCLIBTABLE1,TOT_RTRSCLIBTABLE2,TOT_RTRSCLIBTABLE3]) ;
registerclasses([TOT_RTRSCLIBTABLE4,TOT_RTRSCLIBTABLE5,TOT_RTRSCLIBTABLE6,TOT_RTRSCLIBTABLE7]) ;
registerclasses([TOT_RTRSCLIBTABLE8,TOT_RTRSCLIBTABLE9,TOT_RTMOTIFFERMETURE]) ;
registerclasses([TOT_RTOBJETOPE]) ;
registerclasses([TOT_RT002LIBTABLE0,TOT_RT002LIBTABLE1,TOT_RT002LIBTABLE2]) ;
registerclasses([TOT_RT002LIBTABLE3,TOT_RT002LIBTABLE4]) ;
registerclasses([TOT_RT002LIBTABMUL0,TOT_RT002LIBTABMUL1,TOT_RT002LIBTABMUL2]) ;
registerclasses([TOT_RT002LIBTABMUL3,TOT_RT002LIBTABMUL4]) ;
registerclasses([TOT_RFRPRLIBACTION1,TOT_RFRPRLIBACTION2,TOT_RFRPRLIBACTION3]) ;
registerclasses([TOT_RFLIBCHAINAGE1,TOT_RFLIBCHAINAGE2,TOT_RFLIBCHAINAGE3]) ;
registerclasses([TOT_RFOBJETOPE]) ;
registerclasses([TOT_RTRPRLIBTABLE0,TOT_RTRPRLIBTABLE1,TOT_RTRPRLIBTABLE2]) ;
registerclasses([TOT_RTRPRLIBTABLE3,TOT_RTRPRLIBTABLE4,TOT_RTRPRLIBTABLE5]) ;
registerclasses([TOT_RTRPRLIBTABLE6,TOT_RTRPRLIBTABLE7,TOT_RTRPRLIBTABLE8]) ;
registerclasses([TOT_RTRPRLIBTABLE9,TOT_RTRPRLIBTABLE10,TOT_RTRPRLIBTABLE11]) ;
registerclasses([TOT_RTRPRLIBTABLE12,TOT_RTRPRLIBTABLE13,TOT_RTRPRLIBTABLE14]) ;
registerclasses([TOT_RTRPRLIBTABLE15,TOT_RTRPRLIBTABLE16,TOT_RTRPRLIBTABLE17]) ;
registerclasses([TOT_RTRPRLIBTABLE18,TOT_RTRPRLIBTABLE19,TOT_RTRPRLIBTABLE20]) ;
registerclasses([TOT_RTRPRLIBTABLE21,TOT_RTRPRLIBTABLE22,TOT_RTRPRLIBTABLE23]) ;
registerclasses([TOT_RTRPRLIBTABLE24,TOT_RTRPRLIBTABLE25,TOT_RTRPRLIBTABLE26]) ;
registerclasses([TOT_RTRPRLIBTABLE27,TOT_RTRPRLIBTABLE28,TOT_RTRPRLIBTABLE29]) ;
registerclasses([TOT_RTRPRLIBTABLE30,TOT_RTRPRLIBTABLE31,TOT_RTRPRLIBTABLE32]) ;
registerclasses([TOT_RTRPRLIBTABLE33,TOT_RTRPRLIBTABLE34]) ;
registerclasses([TOT_RTRPRLIBTABMUL0,TOT_RTRPRLIBTABMUL1,TOT_RTRPRLIBTABMUL2]) ;
registerclasses([TOT_RTRPRLIBTABMUL3,TOT_RTRPRLIBTABMUL4,TOT_RTRPRLIBTABMUL5]) ;
registerclasses([TOT_RTRPRLIBTABMUL6,TOT_RTRPRLIBTABMUL7,TOT_RTRPRLIBTABMUL8,TOT_RTRPRLIBTABMUL9]) ;
registerclasses([TOT_RT003LIBTABLE0,TOT_RT003LIBTABLE1,TOT_RT003LIBTABLE2]) ;
registerclasses([TOT_RT003LIBTABLE3,TOT_RT003LIBTABLE4]) ;
registerclasses([TOT_RT003LIBTABMUL0,TOT_RT003LIBTABMUL1,TOT_RT003LIBTABMUL2]) ;
registerclasses([TOT_RT003LIBTABMUL3,TOT_RT003LIBTABMUL4]) ;
registerclasses([TOT_RT006LIBTABLE0,TOT_RT006LIBTABLE1,TOT_RT006LIBTABLE2]) ;
registerclasses([TOT_RT006LIBTABLE3,TOT_RT006LIBTABLE4]) ;
registerclasses([TOT_RT006LIBTABMUL0,TOT_RT006LIBTABMUL1,TOT_RT006LIBTABMUL2]) ;
registerclasses([TOT_RT006LIBTABMUL3,TOT_RT006LIBTABMUL4]) ;
registerclasses([TOT_RTRPCLIBTABLE0,TOT_RTRPCLIBTABLE1,TOT_RTRPCLIBTABLE2]) ;
registerclasses([TOT_RTRPCLIBTABLE3,TOT_RTRPCLIBTABLE4,TOT_RTRPCLIBTABLE5]) ;
registerclasses([TOT_RTRPCLIBTABLE6,TOT_RTRPCLIBTABLE7,TOT_RTRPCLIBTABLE8]) ;
registerclasses([TOT_RTRPCLIBTABLE9,TOT_RTRPCLIBTABLEA,TOT_RTRPCLIBTABLEB]) ;
registerclasses([TOT_RTLIBGED1,TOT_RTLIBGED2,TOT_RTLIBGED3]) ;
{$IFDEF SAV}
registerclasses([TOT_WLIBREWIV1]) ; registerclasses([TOT_WLIBREWIV2]) ;registerclasses([TOT_WLIBREWIV3]) ;
registerclasses([TOT_WLIBREWIV4]) ; registerclasses([TOT_WLIBREWIV5]) ;registerclasses([TOT_WLIBREWIV6]) ;
registerclasses([TOT_WLIBREWIV7]) ; registerclasses([TOT_WLIBREWIV8]) ;registerclasses([TOT_WLIBREWIV9]) ;
registerclasses([TOT_WLIBREWIVA]) ;
registerclasses([TOT_RT007LIBTABLE1]) ; registerclasses([TOT_RT007LIBTABLE2]) ;registerclasses([TOT_RT007LIBTABLE3]) ;
registerclasses([TOT_RT007LIBTABLE4]) ; registerclasses([TOT_RT007LIBTABLE0]) ;
registerclasses([TOT_RT007LIBTABMUL1]) ; registerclasses([TOT_RT007LIBTABMUL2]) ;registerclasses([TOT_RT007LIBTABMUL3]) ;
registerclasses([TOT_RT007LIBTABMUL4]) ; registerclasses([TOT_RT007LIBTABMUL0]) ;
{$ENDIF SAV}
registerclasses([TOT_RT00QLIBTABLE0,TOT_RT00QLIBTABLE1,TOT_RT00QLIBTABLE2]) ;
registerclasses([TOT_RT00QLIBTABLE3,TOT_RT00QLIBTABLE4]) ;
registerclasses([TOT_RT00QLIBTABMUL0,TOT_RT00QLIBTABMUL1,TOT_RT00QLIBTABMUL2]) ;
registerclasses([TOT_RT00QLIBTABMUL3,TOT_RT00QLIBTABMUL4]) ;
registerclasses([TOT_RT00VLIBTABLE0,TOT_RT00VLIBTABLE1,TOT_RT00VLIBTABLE2]) ;
registerclasses([TOT_RT00VLIBTABLE3,TOT_RT00VLIBTABLE4]) ;
registerclasses([TOT_RT00VLIBTABMUL0,TOT_RT00VLIBTABMUL1,TOT_RT00VLIBTABMUL2]) ;
registerclasses([TOT_RT00VLIBTABMUL3,TOT_RT00VLIBTABMUL4]) ;
registerClasses([TOT_RTCORRESPIMPORT]);
registerclasses([TOT_RBCATEGORIE,TOT_RBMULTICHOIXBC1,TOT_RBTYPEINFORMATION]) ;
end.

