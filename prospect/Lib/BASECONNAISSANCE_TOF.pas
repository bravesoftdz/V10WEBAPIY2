{***********UNITE*************************************************
Auteur  ...... : 
Créé le ...... : 20/12/2006
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : BASECONNAISSANCE ()
Mots clefs ... : TOF;BASECONNAISSANCE
*****************************************************************}
Unit BASECONNAISSANCE_TOF ;

Interface

Uses StdCtrls, 
     Controls, 
     Classes, 
{$IFNDEF EAGLCLIENT}
     db,Fe_Main,
     {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF}
     mul,
{$else}
     eMul,MainEAgl,
     uTob,
{$ENDIF}
     forms,
     sysutils,
     ComCtrls,
     HCtrls,
     HEnt1,
     HMsgBox,
     UTOF,UtilBaseConnaissance ;

Type
  TOF_BASECONNAISSANCE = Class (TOF)
    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
  Private
    procedure AfficheChampsLibres ;
    procedure Val1_ElipsisClick ( Sender : tObject);
    procedure Val2_ElipsisClick ( Sender : tObject);
    procedure Val3_ElipsisClick ( Sender : tObject);
    procedure Rattachement1_OnChangeClick ( Sender : tObject);
    procedure Rattachement2_OnChangeClick ( Sender : tObject);
    procedure Rattachement3_OnChangeClick ( Sender : tObject);
    procedure Rattachement_OnChangeClick ( Sender : tObject);
    procedure Val_ElipsisClick ( Sender : tObject);
  end ;

Function RTLanceFiche_BaseConnaissance(Nat,Cod : String ; Range,Lequel,Argument : string) : string;

Implementation
{$IFDEF SAV}
uses entgc;
{$ENDIF SAV}
Function RTLanceFiche_BaseConnaissance(Nat,Cod : String ; Range,Lequel,Argument : string) : string;
begin
Result:=AGLLanceFiche(Nat,Cod,Range,Lequel,Argument);
end;

procedure TOF_BASECONNAISSANCE.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_BASECONNAISSANCE.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_BASECONNAISSANCE.OnUpdate ;
begin
  Inherited ;
end ;

procedure TOF_BASECONNAISSANCE.OnLoad ;
var rattachWhere,ListeCombos,debutWhere,GlobalWhere : string;
    i : integer;
begin
  Inherited ;
  rattachWhere := '';
  GlobalWhere := '';
  debutWhere:=' ( EXISTS ( SELECT LBC_VALEURR FROM LIENBCONNAISSANCE WHERE ( LBC_NUMERO = BCO_NUMERO';

  for i:=1 to 3 do
    begin
    if (GetControlText('RATTACHEMENT'+intToStr(i)) <> '') and (GetControlText('RATTACHEMENT'+intToStr(i)) <> TraduireMemoire('<<Tous>>')) then
      rattachWhere:=debutWhere + ' AND LBC_NUMRATT='+intToStr(i)+' AND LBC_RATTACHEMENT = "'+GetControlText('RATTACHEMENT'+intToStr(i))+'"';

    if (GetControlText('VALEUR'+intToStr(i)) <> '') and (GetControlText('VALEUR'+intToStr(i)) <> TraduireMemoire('<<Tous>>')) then
      begin
      ListeCombos := GetControlText('VALEUR'+intToStr(i));
      if copy(ListeCombos,length(ListeCombos),1) <> ';' then ListeCombos := ListeCombos + ';';
      ListeCombos:=FindEtReplace(ListeCombos,';','","',True);
      ListeCombos:='("'+copy(ListeCombos,1,Length(ListeCombos)-2)+')';
      if rattachWhere = '' then
        rattachWhere:=debutWhere + ' AND LBC_VALEURR in '+ListeCombos
      else
        rattachWhere:=rattachWhere + ' AND LBC_VALEURR in '+ListeCombos;

      if GlobalWhere = '' then
        GlobalWhere:=rattachWhere+')))'
      else
        GlobalWhere:=GlobalWhere+' AND '+rattachWhere+')))';
      end
    else
      if rattachWhere <> '' then
        if GlobalWhere = '' then
          GlobalWhere:=rattachWhere+')))'
        else
          GlobalWhere:=GlobalWhere+' AND '+rattachWhere+')))';
      rattachWhere:='';
    end;

  if (GetControlText('RATTACHEMENT') <> '') and (GetControlText('RATTACHEMENT') <> TraduireMemoire('<<Tous>>')) then
      rattachWhere:=debutWhere + ' AND LBC_RATTACHEMENT = "'+GetControlText('RATTACHEMENT')+'"';
  if (GetControlText('VALEUR') <> '') and (GetControlText('VALEUR') <> TraduireMemoire('<<Tous>>')) then
    begin
    ListeCombos := GetControlText('VALEUR');
    if copy(ListeCombos,length(ListeCombos),1) <> ';' then ListeCombos := ListeCombos + ';';
    ListeCombos:=FindEtReplace(ListeCombos,';','","',True);
    ListeCombos:='("'+copy(ListeCombos,1,Length(ListeCombos)-2)+')';
    if rattachWhere = '' then
      rattachWhere:=debutWhere + ' AND LBC_VALEURR in '+ListeCombos
    else
      rattachWhere:=rattachWhere + ' AND LBC_VALEURR in '+ListeCombos;
    end;
    if rattachWhere <> '' then
      if GlobalWhere = '' then
        GlobalWhere:=rattachWhere+')))'
      else
        GlobalWhere:=GlobalWhere+' AND '+rattachWhere+')))';

  SetControlText('XX_WHERE',GlobalWhere);
end ;

procedure TOF_BASECONNAISSANCE.OnArgument (S : String ) ;
var i : integer;
begin
  Inherited ;
  AfficheChampsLibres;
  if Assigned(GetControl('VALEUR1')) then
    THEdit(GetControl('VALEUR1')).OnElipsisClick := Val1_ElipsisClick;
  if Assigned(GetControl('VALEUR2')) then
    THEdit(GetControl('VALEUR2')).OnElipsisClick := Val2_ElipsisClick;
  if Assigned(GetControl('VALEUR3')) then
    THEdit(GetControl('VALEUR3')).OnElipsisClick := Val3_ElipsisClick;
  if Assigned(GetControl('VALEUR')) then
    THEdit(GetControl('VALEUR')).OnElipsisClick := Val_ElipsisClick;
  if Assigned(GetControl('RATTACHEMENT1')) then
    THVALCOMBOBOX(GetControl('RATTACHEMENT1')).OnChange := Rattachement1_OnChangeClick;
  if Assigned(GetControl('RATTACHEMENT2')) then
    THVALCOMBOBOX(GetControl('RATTACHEMENT2')).OnChange := Rattachement2_OnChangeClick;
  if Assigned(GetControl('RATTACHEMENT3')) then
    THVALCOMBOBOX(GetControl('RATTACHEMENT3')).OnChange := Rattachement3_OnChangeClick;
  if Assigned(GetControl('RATTACHEMENT')) then
    THVALCOMBOBOX(GetControl('RATTACHEMENT')).OnChange := Rattachement_OnChangeClick;
{$IFDEF SAV}
  { si SAV pas sérialisé, on enlève la table du combo }
  if not VH_GC.SAVSeria then
{$ENDIF SAV}
    begin
    for i := 1 to nb_rattachement Do
      if Assigned(GetControl('RATTACHEMENT'+IntToStr(i))) then
        THVALCOMBOBOX(GetControl('RATTACHEMENT'+IntToStr(i))).plus := ' AND CO_CODE <> "'+PrefixeArtParc+'"';
    THVALCOMBOBOX(GetControl('RATTACHEMENT')).plus := ' AND CO_CODE <> "'+PrefixeArtParc+'"';
    end;
end ;

procedure TOF_BASECONNAISSANCE.AfficheChampsLibres ;
var i : integer;
begin
  for i := 1 to 3 Do
    if copy(RechDom('RBBOOLLIBREBC','BB'+IntToStr(i),FALSE),1,1) = '.' then
      SetControlVisible ('BCO_BOOLLIBRE'+IntToStr(i),false)
    else
      SetControlCaption ('BCO_BOOLLIBRE'+IntToStr(i),RechDom('RBBOOLLIBREBC','BB'+IntToStr(i),FALSE));
  SetControlCaption ('TBCO_CHARLIBRE1',RechDom('RBMULTILIBREBC','BM1',FALSE));
end ;

procedure TOF_BASECONNAISSANCE.Val1_ElipsisClick( Sender : tObject);
var retour : string;
begin
  retour:= LanceTheGoodMul(THVALCOMBOBOX(GetControl('RATTACHEMENT1')),nil,nil,nil,1);
  if retour <>'' then
    SetControlText('VALEUR1',retour);
end ;

procedure TOF_BASECONNAISSANCE.Val2_ElipsisClick( Sender : tObject);
var retour : string;
begin
  retour:= LanceTheGoodMul(nil,THVALCOMBOBOX(GetControl('RATTACHEMENT2')),nil,nil,2);
  if retour <>'' then
    SetControlText('VALEUR2',retour);
end ;

procedure TOF_BASECONNAISSANCE.Val3_ElipsisClick( Sender : tObject);
var retour : string;
begin
  retour:= LanceTheGoodMul(nil,nil,THVALCOMBOBOX(GetControl('RATTACHEMENT3')),nil,3);
  if retour <>'' then
    SetControlText('VALEUR3',retour);
end ;

procedure TOF_BASECONNAISSANCE.Val_ElipsisClick( Sender : tObject);
var retour : string;
begin
  retour:= LanceTheGoodMul(nil,nil,nil,THVALCOMBOBOX(GetControl('RATTACHEMENT')),4);
  if retour <>'' then
    SetControlText('VALEUR',retour);
end ;

procedure TOF_BASECONNAISSANCE.Rattachement1_OnChangeClick ( Sender : tObject);
begin
  if GetControlText('RATTACHEMENT1') = '' then
    SetControlText('VALEUR1','');
end;

procedure TOF_BASECONNAISSANCE.Rattachement2_OnChangeClick ( Sender : tObject);
begin
  if GetControlText('RATTACHEMENT2') = '' then
    SetControlText('VALEUR2','');
end;

procedure TOF_BASECONNAISSANCE.Rattachement3_OnChangeClick ( Sender : tObject);
begin
  if GetControlText('RATTACHEMENT3') = '' then
    SetControlText('VALEUR3','');
end;

procedure TOF_BASECONNAISSANCE.Rattachement_OnChangeClick(Sender : tObject);
begin
  if GetControlText('RATTACHEMENT') = '' then
    SetControlText('VALEUR','');
end;

Initialization
  registerclasses ( [ TOF_BASECONNAISSANCE ] ) ;
end.

