{***********UNITE*************************************************
Auteur  ...... :
Créé le ...... : 11/01/2001
Modifié le ... :   /  /
Description .. : Source TOM de la TABLE : CODEBARRES (CODEBARRES)
Mots clefs ... : TOM;CODEBARRES
*****************************************************************}
Unit UTomCodeBarres ;

Interface

Uses StdCtrls, Controls, Classes, db, forms, sysutils, dbTables, ComCtrls,
     HCtrls, HEnt1, HMsgBox, UTOM, Fiche, FichList, UTob ;

Type
  TOM_CodeBarres = Class (TOM)
  private
    Function PrefixeCABOk : boolean;
    Function TypeCABOk : boolean;
    Function LgNumCABOk : boolean;
    Procedure ActiveChamp ;
    Function ControlePrefixe ( Prefixe : String; CaractereAutorise : String ) : boolean ;
    Function ControleChrono ( Chrono : String ) : boolean ;
    Function ControleLgCAB ( LgMin:integer;LgMax:integer;LgCode:integer;LgPrefixe:integer;LgCompteur:integer;TypeCAB:String) : boolean;
    Procedure SetLastError (Num : integer; ou : string );
  public
    procedure OnNewRecord                ; override ;
    procedure OnDeleteRecord             ; override ;
    procedure OnUpdateRecord             ; override ;
    procedure OnAfterUpdateRecord        ; override ;
    procedure OnLoadRecord               ; override ;
    procedure OnChangeField ( F: TField) ; override ;
    procedure OnArgument ( S: String )   ; override ;
    procedure OnClose                    ; override ;
    procedure OnCancelRecord             ; override ;
    end ;

Implementation

uses ParamSoc ;

const
	// libellés des messages
	TexteMessage: array[1..9] of string 	= (
          {1}        'Le préfixe doit être numérique'
          {2}        ,'Code préfixe incorrect'
          {3}        ,'Le chrono doit être numérique'
          {4}        ,'Longueur minimum du code : '
          {5}        ,'Longueur maximum du code : '
          {6}        ,'La longueur du chrono est trop grande'
          {7}        ,'La longueur du code est trop petite'
          {8}        ,'La longueur du code doit être un nombre pair'
          {9}        ,'Type de code à barres incorrect'
                     );

procedure TOM_CodeBarres.SetLastError (Num : integer; ou : string );
begin
if ou<>'' then SetFocusControl(ou);
LastError:=Num;
LastErrorMsg:=TexteMessage[LastError];
end ;

procedure TOM_CodeBarres.OnNewRecord ;
begin
  Inherited ;
   if (GetField('GCB_NATURECAB')='SO')
     then SetField('GCB_CABTIERS', 'X')
     else SetField('GCB_CABTIERS', '-');
   SetField('GCB_NUMCABAUTO', 'X');
   SetControlText('NUMCABAUTO', 'X');
   SetField('GCB_TYPECAB', 'E13');
   SetField('GCB_PREFIXECAB', '');
   SetField('GCB_LGNUMCAB', 13);
   SetField('GCB_COMPTEURCAB', '0');
end ;

procedure TOM_CodeBarres.OnDeleteRecord ;
begin
  Inherited ;
end ;

Function TOM_CodeBarres.PrefixeCABOk : boolean;
Var Prefixe : String;
    TypeCAB : String;
const
    CaractereAutorise : String ='0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZ-. $/+%' ;
begin
     Result:=True;
     Prefixe:=Trim(GetField('GCB_PREFIXECAB'));
     SetControlText('GCB_PREFIXECAB', Prefixe);
     if Prefixe <> '' then
        begin
        TypeCAB:=GetField('GCB_TYPECAB');
        if ((TypeCAB='E13') or (TypeCAB='EA8') or (TypeCAB='ITC') or (TypeCAB='ITF'))
            and (not IsNumeric(Prefixe)) then
           begin
           SetLastError(1, 'GCB_PREFIXECAB');
           Result:=False;
           Exit;
           end;
        if ControlePrefixe (Prefixe, CaractereAutorise)=False then
           begin
           SetLastError(2, 'GCB_PREFIXECAB');
           Result:=False;
           end;
        end;
end ;

Function TOM_CodeBarres.TypeCABOk : boolean;
Var TypeCAB : String;
begin
     Result:=True;
     TypeCAB:=GetField('GCB_TYPECAB');
     if GetControlText('NUMCABAUTO')='X' then
        begin
        if (TypeCAB='') or (TypeCAB='128') then
           begin
           SetLastError(9, 'GCB_TYPECAB');
           Result:=False;
           Exit;
           end;
        end;
     if (TypeCAB='EA8') then SetField('GCB_LGNUMCAB', 8);
     if (TypeCAB='E13') then SetField('GCB_LGNUMCAB', 13);
end ;

Function TOM_CodeBarres.LgNumCABOk : boolean;
Var CompteurCAB,TypeCAB : String;
    LgCompteur, LgPrefixe, LgNumCAB : Integer;
begin
     Result:=True;
     LgPrefixe := Length(GetField('GCB_PREFIXECAB'));
     LgNumCAB := GetField('GCB_LGNUMCAB');
     TypeCAB := GetField('GCB_TYPECAB');
     CompteurCAB := Trim(GetField('GCB_COMPTEURCAB'));
     if CompteurCAB='' then
        begin
        CompteurCAB:='0';
        SetControlText('GCB_COMPTEURCAB', CompteurCAB);
        end;
     LgCompteur := Length(CompteurCAB);
     if ControleChrono (CompteurCAB)=False then
        begin   // Contrôle de numéricité
        Result:=False;
        Exit;
        end;
     if ControleLgCAB ( 4, 18, LgNumCAB, LgPrefixe, LgCompteur, TypeCAB)=False then
        begin
        Result:=False;
        Exit;
        end;
end ;

procedure TOM_CodeBarres.ActiveChamp ;
Var Active,ActiveLg : boolean;
begin
   if (GetField('GCB_NATURECAB')='SO') then
      begin
      // Paramétrage global du code à barres
      SetControlVisible('GCB_CABTIERS', False);
      SetControlEnabled('NUMCABAUTO', True);
      end else
      begin
      // Paramétrage du code à barres par fournisseur
      if (GetField('GCB_CABTIERS')='X') then Active:=True else Active:=False ;
      if GetParamsoc('SO_GCNUMCABAUTO')=False then Active:=False;
      SetControlEnabled('NUMCABAUTO', Active);
      end;

   if (GetField('GCB_TYPECAB')='E13') or (GetField('GCB_TYPECAB')='EA8')
      then ActiveLg:=False else ActiveLg:=True;

   if (GetField('GCB_CABTIERS')='X') and (GetControlText('NUMCABAUTO')='X')
      then Active:=True else Active:=False;

   SetControlEnabled('GCB_PREFIXECAB',Active);
   SetControlEnabled('TGCB_PREFIXECAB',Active);
   SetControlEnabled('GCB_LGNUMCAB',ActiveLg);
   SetControlEnabled('TGCB_LGNUMCAB',ActiveLg);
   SetControlEnabled('GCB_COMPTEURCAB',Active);
   SetControlEnabled('TGCB_COMPTEURCAB',Active);

   if GetControlText ('NUMCABAUTO')='X' then
      begin
      SetControlProperty('GCB_TYPECAB','Vide', False);
      SetControlProperty('GCB_TYPECAB','Plus', ' AND (CO_CODE="39" OR CO_CODE="39C" OR CO_CODE="E13" OR CO_CODE="EA8" OR CO_CODE="ITC" OR CO_CODE="ITF")') ;
      end else
      begin
      SetControlProperty('GCB_TYPECAB','Vide', True);
      SetControlProperty('GCB_TYPECAB','Plus', '');
      end;
end ;

Function TOM_CodeBarres.ControlePrefixe ( Prefixe : String; CaractereAutorise : String ) : boolean ;
Var i, j : Integer;
    CarOk: Boolean;
BEGIN
Result:=True;
for i:=1 to Length(Prefixe) do
    begin
    CarOk:=False;
    for j:=1 to Length(CaractereAutorise) do
        begin
        if Prefixe[i]=CaractereAutorise[j] then begin CarOk:=True; break; end;
        end;
    if CarOk = False then begin Result:=False; break; end;
    end;
END ;

Function TOM_CodeBarres.ControleChrono ( Chrono : String ) : boolean ;
BEGIN
Result:=True;
if (not IsNumeric(Chrono))then
   begin
   SetLastError(3, 'GCB_COMPTEURCAB');
   Result:=False;
   end;
END ;

Function TOM_CodeBarres.ControleLgCAB ( LgMin:integer;LgMax:integer;LgCode:integer;LgPrefixe:integer;LgCompteur:integer;TypeCAB:String) : boolean;
Var LgCle : Integer;
    ChampEcran : String;
BEGIN
Result:=True;
if (TypeCAB='E13') or (TypeCAB='EA8') or (TypeCAB='39C') or (TypeCAB='ITC') or (TypeCAB='ITF')
   then LgCle:=1 else LgCle:=0;
if (TypeCAB='E13') then LgMax:=13;
if (TypeCAB='EA8') then LgMax:=8;
if (TypeCAB='E13') or (TypeCAB='EA8')
   then ChampEcran:='GCB_COMPTEURCAB'
   else ChampEcran:='GCB_LGNUMCAB';
if (TypeCAB='ITC') or (TypeCAB='ITF') then
   begin
   // la longueur du code 2/5 Entrelacé doit être de longueur paire
   if (LgCode mod 2)<>0 then
      begin
      SetLastError(8, ChampEcran);
      Result := False;
      Exit;
      end;
   end;
if (LgCode<LgPrefixe+LgMin) then
   begin
   SetLastError(4, ChampEcran);
   LastErrorMsg:=LastErrorMsg+InttoStr(LgPrefixe+LgMin);
   Result := False;
   Exit;
   end;
if (LgCode>LgMax) then
   begin
   SetLastError(5, ChampEcran);
   LastErrorMsg:=LastErrorMsg+InttoStr(LgMax);
   Result := False;
   Exit;
   end;
if (LgCompteur > (LgCode-LgPrefixe-LgCle)) then
   begin
   if LgCle <> 0 then SetLastError(6, ChampEcran)
                 else SetLastError(7, ChampEcran) ;
   Result := False;
   Exit;
   end;
END ;

procedure TOM_CodeBarres.OnUpdateRecord ;
begin
  Inherited ;
  If PrefixeCABOk=False then Exit;
  If TypeCABOk=False then Exit;
  If LgNumCABOk=False then Exit;
  SetField('GCB_NUMCABAUTO', GetControlText('NUMCABAUTO'));
  if (GetField('GCB_NATURECAB')='SO') then SetParamSoc('SO_GCNUMCABAUTO', GetControlText('NUMCABAUTO'));
end ;

procedure TOM_CodeBarres.OnAfterUpdateRecord ;
begin
  Inherited ;
end ;


procedure TOM_CodeBarres.OnLoadRecord ;
begin
  Inherited ;
   if (GetField('GCB_NATURECAB')='SO') then
      begin
      TForm(Ecran).Caption := 'Code à barres article :';
      SetControlChecked('NUMCABAUTO',GetParamSoc('SO_GCNUMCABAUTO')) ;
      SetFocusControl('GCB_TYPECAB');
      end else
      begin
      if GetParamSoc ('SO_GCNUMCABAUTO')=True
         then SetControlText('NUMCABAUTO', GetField('GCB_NUMCABAUTO'))
         else SetControlChecked('NUMCABAUTO',False) ;
      SetFocusControl('GCB_CABTIERS');
      end;
end ;

procedure TOM_CodeBarres.OnChangeField ( F: TField ) ;
begin
  Inherited ;
   If (F.FieldName='GCB_PREFIXECAB') and (PrefixeCABOk=False) then Exit;

   If (F.FieldName='GCB_TYPECAB') and (TypeCABOk=False) then Exit;

   If (F.FieldName='GCB_CABTIERS') or (F.FieldName='GCB_NUMCABAUTO') or
      (F.FieldName='GCB_TYPECAB') then ActiveChamp;
end ;

procedure TOM_CodeBarres.OnArgument ( S: String ) ;
begin
  Inherited ;
end ;

procedure TOM_CodeBarres.OnClose ;
begin
  Inherited ;
end ;

procedure TOM_CodeBarres.OnCancelRecord ;
begin
  Inherited ;
end ;

Initialization
  registerclasses ( [ TOM_CodeBarres ] ) ;
end.
