unit TofGLage;

interface

uses Classes, StdCtrls,
{$IFDEF EAGLCLIENT}
     eQRS1,MaineAGL,
{$ELSE}
     dbtables,QRS1,FE_Main,
{$ENDIF}
     comctrls,UTof, HCtrls,Ent1,  TofMeth, HTB97, spin,
     SysUtils;

type
    TOF_GLAGE = class(TOF_Meth)
  private
    Periodicite     : THValComboBox ;
    Etablissement   : THValComboBox ;
    Devise          : THValComboBox ;
    NatureAuxi      : THValComboBox ;
    FP1             : THEdit ;
    FP8             : THEdit ;
    DevPivot        : THEdit ;
    Ecart           : THRadioGroup ;
    Affichage       : THRadioGroup ;
    Impression      : THRadioGroup ;
    NbJEcart        : TSpinEdit ;
    ChargeFiltre    : Boolean ;
    Reference       : THValComboBox ;
    Libelle         : THValComboBox ;
    Aux1, Aux2      : THEdit ;
    Gen1, Gen2      : THEdit ;
    procedure NatureAuxiOnChange (Sender : TObject) ;
    procedure ZFiltreOnChange (Sender : TObject) ;
    procedure PeriodiciteOnChange (Sender : TObject) ;
    procedure DateArreteOnChange (Sender : TObject) ;
    procedure EcartOnChange (Sender : TObject) ;
    procedure NbJEcartOnChange(Sender : TObject);
    procedure TypeEcrNOnClick (Sender : TObject);
    procedure TypeEcrSOnClick (Sender : TObject);
    procedure TypeEcrUOnClick (Sender : TObject);
    procedure EtablissementOnChange (Sender : TObject);
    procedure DeviseOnChange (Sender : TObject);
    procedure Rupture1OnDblClick (Sender : TObject);
    procedure TriParOnDblClick (Sender : TObject);
    procedure AffichageOnClick (Sender : TObject);
    procedure ImpressionOnClick (Sender : TObject);
    procedure TriParOnChange (Sender : TObject);
    procedure ReferenceOnChange (Sender : TObject);
    procedure LibelleOnChange (Sender : TObject);

  public
    procedure OnArgument(Arguments : string); override ;
    procedure OnUpdate; override ;
  end ;

Type TTabDate8 = Array[1..8] of TDateTime ;

implementation

uses HEnt1, LicUtil,HMsgBox;

PROCEDURE TOF_GLAGE.OnUpdate;
VAR St,St1,St2,St3,St4,St5,S,S1,TriLibelle,Order : string;
    i, j : integer;
BEGIN
S:='' ; S1:='' ; j:=0 ;
TriLibelle:='E_AUXILIAIRE';
St:=GetControlText('TRIPAR');
St1:=GetControlText('RUPTURE1');
St2:=GetControlText('RUPTURE2');
if GetControlText('TRILIBELLE')='X' then TriLibelle:='T_LIBELLE' else TriLibelle:='E_AUXILIAIRE';
While St1<>'' do
  For i:=0 to 3 do
    BEGIN
    St3:=ReadTokenSt(St1) ;
    St4:=ReadTokenSt(St2) ;
    if (St3<>'-') and (St3>'#') and (St3<>'') then
    BEGIN
    if (St3<>'*') then
      if TCheckBox(GetControl('CPTASSOCIES')).Checked=FALSE then
        BEGIN
        if j<1 then S:=S+' AND (E_TABLE'+IntToStr(i)+'>="'+St3+'" AND E_TABLE'+IntToStr(i)+'<="'+St4+'") '
        else S:=S+' OR (E_TABLE'+IntToStr(i)+'>="'+St3+'" AND E_TABLE'+IntToStr(i)+'<="'+St4+'") ' ;
        j:=j+1 ;
        END else S:=S+' AND (E_TABLE'+IntToStr(i)+'>="'+St3+'" AND E_TABLE'+IntToStr(i)+'<="'+St4+'") ' ;
    if (St3='*') then
      if TCheckBox(GetControl('CPTASSOCIES')).Checked=TRUE then
        BEGIN
        S:=S+' AND (E_TABLE'+IntToStr(i)+'<>"") ' ;
        j:=j+1 ;
        END ;
    END ;
  END ;

while St<>'' do
  for i:=0 to 3 do
  BEGIN
  St5:=ReadTokenSt(St);
  if (St5<>'') then
    BEGIN
    SetControlText('ORDER'+IntToStr(i),'E_TABLE'+copy(St5,3,1));
    SetControlText('E0'+IntToStr(i),St5);
    if (St='')then S1:=S1+'E_TABLE'+copy(St5,3,1) else S1:=S1+'E_TABLE'+copy(St5,3,1)+',';
    END else
    BEGIN
    SetControlText('E0'+IntToStr(i),'');
    SetControlText('ORDER'+IntToStr(i),'');
    END ;
  END;

if (Aux1.Text<>'') and (Aux2.Text='')  then S:=S+' AND E_AUXILIAIRE>="'+Aux1.Text+'"' ;
if (Aux1.Text='')  and (Aux2.Text<>'') then S:=S+' AND E_AUXILIAIRE<="'+Aux2.Text+'"' ;
if (Aux1.Text<>'') and (Aux2.Text<>'') then S:=S+' AND E_AUXILIAIRE>="'+Aux1.Text+'" AND E_AUXILIAIRE<="'+Aux2.Text+'"';
if NatureAuxi.Itemindex<>0             then S:=S+' AND T_NATUREAUXI="'+NatureAuxi.Value+'"';
if (Gen1.Text<>'') and (Gen2.Text='')  then S:=S+' AND E_GENERAL>="'+Gen1.Text+'"';
if (Gen1.Text='')  and (Gen2.Text<>'') then S:=S+' AND E_GENERAL<="'+Gen2.Text+'"';
if (Gen1.text<>'') and (Gen2.Text<>'') then S:=S+' AND E_GENERAL>="'+Gen1.text+'" AND E_GENERAL<="'+Gen2.Text+'"';
if S<>''  then SetControlText('WHERE',S);
if S1<>'' then Order:=S1+','+TriLibelle+',E_DATECOMPTABLE' else Order:=TriLibelle+',E_DATECOMPTABLE';
TFQRS1(Ecran).WhereSQL:='|'+Order;
END;

PROCEDURE TOF_GLAGE.OnArgument(Arguments : string);
VAR DateArrete                      : THEdit;
    Rupture1,Rupture2               : THEdit;
    TriPar                          : THEdit;
    TypeEcrN, TypeEcrS, TypeEcrU    : TCheckBox;
    ZFiltre                         : TcomboBox;
BEGIN
ChargeFiltre:=FALSE ;
inherited ;
NatureAuxi:=THValComboBox(GetControl('NATUREAUXI'));
Periodicite:=THValComboBox(GetControl('PERIODICITE'));
Devise:=THValComboBox(GetControl('DEVISE'));
DevPivot:=THEdit(GetControlText('DEVPIVOT'));
Etablissement:=THValComboBox(GetControl('ETABLISSEMENT'));
TypeEcrN:=TCheckBox(GetControl('NORMAL'));
TypeEcrS:=TCheckBox(GetControl('SIMULATION'));
TypeEcrU:=TCheckBox(GetControl('SITUATION'));
FP1:=THEdit(GetControl('PERIODE1'));
FP8:=THEdit(GetControl('DATE1'));
DateArrete:=THEdit(GetControl('DATE1'));
Ecart:=THRadioGroup(GetControl('ECART'));
NbJEcart:=TSpinEdit(GetControl('NBJECART'));
ZFiltre:=TComboBox(GetControl('FFILTRES'));
Rupture1:=THEdit(GetControl('RUPTURE1'));
Rupture2:=THEdit(GetControl('RUPTURE2'));
TriPar:=THEdit(GetControl('TRIPAR'));
Affichage:=THRadioGroup(GetControl('AFFICHAGE'));
Impression:=THRadioGroup(GetControl('TIMPRESSION'));
Reference:=THValComboBox(GetControl('REFERENCE'));
Libelle:=THValComboBox(GetControl('LIBELLE'));
Aux1:=THEdit(GetControl('E_AUXILIAIRE')) ;
Aux2:=THEdit(GetControl('E_AUXILIAIRE_')) ;
Gen1:=THEdit(GetControl('GENERAL')) ;
Gen2:=THEdit(GetControl('GENERAL_')) ;
if GetControlText('NORMAL')    ='X' then SetControlText('N','N') else SetControlText('N','W');
if GetControlText('SIMULATION')='X' then SetControlText('S','S') else SetControlText('S','Y');
if GetControlText('SITUATION') ='X' then SetControlText('U','U') else SetControlText('U','Z');
TypeEcrN.OnClick:=TypeEcrNOnClick;
TypeEcrS.OnClick:=TypeEcrSOnClick;
TypeEcrU.OnClick:=TypeEcrUOnClick;
if Reference<>nil     then BEGIN Reference.OnChange:=ReferenceOnChange ; THValComboBox(GetControl('REFERENCE')).ItemIndex:=0;ReferenceOnChange(nil) ; END ;
if Libelle<>nil       then BEGIN Libelle.OnChange:=LibelleOnChange ; THValComboBox(GetControl('LIBELLE')).ItemIndex:=2 ; LibelleOnChange(nil) ; END ;
if Devise<>nil        then BEGIN Devise.OnChange:=DeviseOnChange ; Devise.itemindex:=0 ; Devise.OnChange(nil) ; END ;
if Etablissement<>nil then BEGIN Etablissement.OnChange:=EtablissementOnChange ; Etablissement.itemindex:=0 ; Etablissement.OnChange(Nil) ; END ;
if DateArrete<>nil    then BEGIN DateArrete.OnChange:=DateArreteOnChange ; SetControlText('DATE1',DateToStr(V_PGI.DateEntree)) ; END ;
if Ecart<>nil         then BEGIN Ecart.OnClick:=EcartOnChange ; Ecart.itemindex:=0;EcartOnChange(nil) ; END ;
if Periodicite<>nil   then BEGIN Periodicite.OnChange:=PeriodiciteOnChange ; Periodicite.itemindex:=0 ; END ;
if NbJEcart<>nil      then BEGIN NbJEcart.OnChange:=NbJEcartOnChange;NbJEcart.text:='30' ; NbJEcartOnChange(nil) ; END ;
if NatureAuxi<>nil    then BEGIN NatureAuxi.OnChange:=NatureAuxiOnChange ; NatureAuxi.itemindex:=2 ; SetControlText('NATUREAUX','CLI') ; END ;
if Affichage<>nil then
  BEGIN
  Affichage.OnClick:=AffichageOnClick;
  Affichage.itemindex:=0;
  SetControlText('DEBIT','E_DEBIT');
  SetControlText('CREDIT','E_CREDIT');
  END ;
if Impression<>nil then
  BEGIN
  Impression.OnClick:=ImpressionOnClick;
  Impression.itemindex:=0;
  SetControlText('IMPRESSION','DECI');
  END ;
if Rupture1<>nil then Rupture1.OnDblClick:=Rupture1OnDblClick;
if Rupture2<>nil then Rupture2.OnDblClick:=Rupture1OnDblClick;
if TriPar<>nil   then
  BEGIN
  TriPar.OnChange:=TriParOnChange;
  TriPar.OnDblClick:=TriParOnDblClick;
  SetControlEnabled('SAUTPAGE',False);
  SetControlEnabled('SURRUPTURE',False);
  END ;
ZFiltre.Onchange:=ZFiltreOnChange ;
END;

PROCEDURE TOF_GLAGE.Rupture1OnDblClick (Sender : TObject);
VAR    St,St1,St2,St3,Arg:String;
       i : integer;
       OkAssoc : integer;
BEGIN
OkAssoc:=0;
Arg:=GetControlText('RUPTURE1')+'|'+GetControlText('RUPTURE2');
St:=AglLanceFiche('CP','TLMVTRUPT','','',Arg);
St1 :=ReadTokenPipe(St,'|');
SetControlText('RUPTURE1',St1);
SetControlText('RUPTURE2',St);
for i:=0 to 3 do
  BEGIN
  St2:=ReadTokenSt(St1);
  if (St2<>'#') and (St2<>'-') and (St2<>'') then BEGIN St3:=St3+'E0'+IntToStr(i)+';'; OkAssoc:=OkAssoc+1 ; END ;
  END ;
if OkAssoc=0 then
  BEGIN
  SetControlEnabled('CPTASSOCIES',False);
  SetControlChecked('CPTASSOCIES',False);
  END else
  BEGIN
  SetControlEnabled('CPTASSOCIES',True);
  SetControlChecked('CPTASSOCIES',True);
  END ;
SetControlText('TRIPAR',St3);
END;

PROCEDURE TOF_GLAGE.TriParOnChange (Sender : TObject);
BEGIN
if GetControlText('TRIPAR')='' then
  BEGIN
  SetControlChecked('SAUTPAGE',False);
  SetControlChecked('SURRUPTURE',False);
  SetControlEnabled('SAUTPAGE',False);
  SetControlEnabled('SURRUPTURE',False);
  END else
  BEGIN
  SetControlEnabled('SAUTPAGE',True);
  SetControlEnabled('SURRUPTURE',True);
  END ;
END ;

PROCEDURE TOF_GLAGE.TriParOnDblClick (Sender : TObject);
VAR    St,Arg : String ;
BEGIN
Arg:=GetControlText('TRIPAR');
St:=AglLanceFiche('CP','TLMVTTRI','','',Arg);
SetControlText('TRIPAR',St);
END;

PROCEDURE TOF_GLAGE.ImpressionOnClick (Sender : TObject);
BEGIN
case Impression.ItemIndex of
  0 : SetControlText('IMPRESSION','DECI') ;
  1 : SetControlText('IMPRESSION','KILO') ;
  2 : SetControlText('IMPRESSION','MEGA') ;
  3 : SetControlText('IMPRESSION','ENTI') ;
  END ;
END;

PROCEDURE TOF_GLAGE.AffichageOnClick (Sender : TObject);
BEGIN
if (Affichage.ItemIndex=2) and ((GetControlText('DEVISE')=V_PGI.DevisePivot) or (THValComboBox(GetControl('DEVISE')).ItemIndex=0)) then
  BEGIN
  PgiBox('Vous devez d''abord sélectionner une devise particulière','Grand-Livre Agé');
  THRadiogroup(GetControl('AFFICHAGE')).ItemIndex:=0;
  END ;
case Affichage.ItemIndex of
  0 : BEGIN SetControlText('DEBIT','E_DEBIT')    ;SetControlText('CREDIT','E_CREDIT')    ; END ;
  1 : BEGIN SetControlText('DEBIT','E_DEBITEURO');SetControlText('CREDIT','E_CREDITEURO'); END ;
  2 : BEGIN SetControlText('DEBIT','E_DEBITDEV') ;SetControlText('CREDIT','E_CREDITDEV') ; END ;
  END ;
END;

PROCEDURE TOF_GLAGE.DeviseOnChange (Sender : TObject);
BEGIN
if Devise.ItemIndex=0               then SetControlText('DEVPIVOT','%')
                                    else SetControlText('DEVPIVOT',GetControlText('DEVISE'));
if (Devise.value=V_PGI.DevisePivot) and (THRadioGroup(GetControl('AFFICHAGE')).ItemIndex=2) then
  BEGIN
  PgiBox('Cette devise n''est pas compatible avec l''affichage sélectionné','Grand-Livre Agé');
  THRadiogroup(GetControl('AFFICHAGE')).ItemIndex:=0;
  END ;
if (THValComboBox(GetControl('DEVISE')).ItemIndex=0) and (THRadioGroup(GetControl('AFFICHAGE')).ItemIndex=2) then
  BEGIN
//  PgiBox('Ce choix n''est pas compatible avec l''affichage sélectionné','Balance Agée');
  THRadiogroup(GetControl('AFFICHAGE')).ItemIndex:=0;
  END ;
END;

PROCEDURE TOF_GLAGE.ReferenceOnChange (Sender : TObject);
BEGIN
SetControlText('LIB1',THValComboBox(GetControl('REFERENCE')).Value);
END;

PROCEDURE TOF_GLAGE.LibelleOnChange (Sender : TObject);
BEGIN
SetControlText('LIB2',THValComboBox(GetControl('LIBELLE')).Value);
END;

PROCEDURE TOF_GLAGE.EtablissementOnChange (Sender : TObject);
BEGIN
if Etablissement.ItemIndex=0 then SetControlText('ETAB','%') else SetControlText('ETAB',GetControlText('ETABLISSEMENT'));
END;

PROCEDURE TOF_GLAGE.TypeEcrNOnClick (Sender : TObject);
BEGIN
if GetControlText('NORMAL')='X' then SetControlText('N','N') else SetControlText('N','W');
END;

PROCEDURE TOF_GLAGE.TypeEcrSOnClick (Sender : TObject);
BEGIN
if GetControlText('SIMULATION')='X' then SetControlText('S','S') else SetControlText('S','Y');
END;

PROCEDURE TOF_GLAGE.TypeEcrUOnClick (Sender : TObject);
BEGIN
if GetControlText('SITUATION')='X' then SetControlText('U','U') else SetControlText('U','Z');
END;

PROCEDURE TOF_GLAGE.ZFiltreOnChange (Sender : TObject) ;
BEGIN
ChargeFiltre:=TRUE ;
inherited ;
ChargeFiltre:=FALSE ;
END ;

PROCEDURE TOF_GLAGE.NatureAuxiOnChange (Sender : TObject) ;
Var CGen, CAux : string ;
BEGIN
inherited ;
If ChargeFiltre Then Exit ;
SetControlText('GENERAL','');
SetControlText('GENERAL','');
SetControlText('NATUREAUX',GetControlText('T_NATUREAUXI'));
Case NatureAuxi.ItemIndex of
  2:  BEGIN CGen:='TZGCOLLCLIENT' ; CAux:='TZTCLIENT' ; END ;
  4:  BEGIN CGen:='TZGCOLLFOURN'  ; CAux:='TZTFOURN'  ; END ;
 else BEGIN CGen:='TZGTOUS'       ; CAux:='TZTTOUS'   ; END ;
END ;
SetControlProperty('E_AUXILIAIRE','DATATYPE',CAux);
SetControlProperty('E_AUXILIAIRE_','DATATYPE',CAux);
SetControlProperty('GENERAL','DATATYPE',CGen);
SetControlProperty('GENERAL_','DATATYPE',CGen);
END ;

PROCEDURE PeriodiciteChange(Iperiodicite : Integer ; FP8,FP1 : THEdit ; var TabD : TTabDate8) ;
VAR Choix,i : Integer ;
    a,m,j, JMax : Word ;
    DAT : TDATETIME ;
BEGIN
//Choix:=FPeriodicite.ItemIndex ;
Choix:=IPeriodicite ;
If Choix<=-1 Then Choix:=0 ;
   TabD[8]:=StrToDate(FP8.text) ;       { Calcul à partir de la date d'arrêtée }
   Case IPeriodicite of
     0,1,2,3,4,5 : for i:=4 downto 1 Do { Pour chaque Fourchette de dates( en partant de la derniére) , en mensuel }
                       BEGIN
                       DAT:=PlusMois(TabD[i+4],-(Choix+1)) ;
                       DecodeDate(DAT,a,m,j) ;
                       JMax:=StrToInt(FormatDateTime('d',FinDeMois(EncodeDate(a,m,1)))) ;
                       TabD[i]:=PlusMois(TabD[i+4],-(Choix+1))+(JMax-J)+1 ;
                       if i>1 then TabD[i+3]:=TabD[i]-1 ;
                       END ;
     6           : for i:=4 downto 1 Do { Pour chaque Fourchette de dates( en partant de la derniére) , en Quinzaine }
                       BEGIN
                       DAT:=TabD[i+4] ;
                       DecodeDate(DAT,a,m,j) ;
                       JMax:=StrToInt(FormatDateTime('d',FinDeMois(EncodeDate(a,m,1)))) ;
                       If J<=15 then
                          BEGIN         { Date départ = (Date d'arrivée - 15 jours) + 1 jour, si date avant le 15 du mois }
                          TabD[i]:=TabD[i+4]-(15  )+1 ;
                          END Else
                          BEGIN         { Date départ = (Date d'arrivée - (Nb jours Max du mois - 15 jours) ) + 1 jour, si date aprés le 15 du mois }
                          TabD[i]:=TabD[i+4]-(JMax-15)+1 ;
                          END ;
                       if i>1 then TabD[i+3]:=TabD[i]-1 ;
                       END ;
     7           : for i:=4 downto 1 Do { Pour chaque Fourchette de dates( en partant de la derniére) , en Hebdo }
                       BEGIN            { Date départ = (Date d'arrivée - (7 jours+ 1 jours) )  }
                       TabD[i]:=TabD[i+4]-(7)+1 ;
                       if i>1 then TabD[i+3]:=TabD[i]-1 ;
                       END ;
     End ;
END;

PROCEDURE TOF_GLAGE.PeriodiciteOnChange(Sender : TObject);
VAR TabD : TTabDate8; I:integer;
BEGIN
inherited ;
FillChar(TabD,SizeOf(TabD),#0);
PeriodiciteChange(Periodicite.itemindex,FP8,FP1, TabD);
for i:=1 to 8 do begin SetControlText('PERIODE'+IntToStr(i),DateToStr(TabD[i])); end;
END;

PROCEDURE TOF_GLAGE.DateArreteOnChange(Sender : TObject);
BEGIN
inherited ;
if not IsValidDate(GetControlText('DATE1')) then exit;
SetControlText('PERIODE8',DateToStr((StrToDate(GetControlText('DATE1')))-1));
EcartOnChange(nil);
END;

PROCEDURE TOF_GLAGE.EcartOnChange(Sender : TObject);
BEGIN
inherited ;
SetControlVisible('PERIODES',Ecart.ItemIndex=1);
if Ecart.ItemIndex=0 then NbJEcartOnChange(nil) else PeriodiciteOnChange(nil);
END;

PROCEDURE NbJEcartCalculDate(FP8 : THEdit ; NbjEcart : integer ; var TabD : TTabDate8) ;
VAR i : Integer ;
BEGIN
TabD[8]:=StrToDate(FP8.Text) ;       { Calcul à partir de la date d'arrêtée }
for i:=4 downto 1 do
  BEGIN
    TabD[i]:=TabD[i+4]-(NbJEcart-1) ;
    if i>1 then TabD[i+3]:=TabD[i]-1 ;
  END ;
END;

PROCEDURE TOF_GLAGE.NbJEcartOnChange(Sender : TObject);
VAR TabD : TTabDate8; i:integer;
BEGIN
inherited;
if GetControlText('NBJECART')='' then exit;
FillChar(TabD,SizeOf(TabD),#0);
NbJEcartCalculDate(FP8,NbJEcart.value, TabD);
for i:=1 to 8 do SetControlText('PERIODE'+IntToStr(i),DateToStr(TabD[i]));
END;

INITIALIZATION
RegisterClasses([TOF_GLAGE]) ;
END.

