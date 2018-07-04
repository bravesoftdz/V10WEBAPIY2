{***********UNITE*************************************************
Auteur  ...... : Régis ROHAULT
Créé le ...... : 26/11/2001
Modifié le ... : 05/12/2001
Description .. : Source TOM de la TABLE : COTATIONTAUX
               : JP 03/08/03 : Migration eAGL
Suite ........ : (COTATIONTAUX)
Mots clefs ... : TOM;COTATIONTAUX
*****************************************************************}
Unit TomCotationTaux ;

Interface

Uses
  {$IFDEF EAGLCLIENT}
  //MaineAGL,
  {$ELSE}
  //FE_Main,
  {$ENDIF}
  Forms, Classes, HCtrls, HTB97, UTOM, SysUtils, Controls;

Type
  TOM_COTATIONTAUX = Class (TOM)
    procedure OnNewRecord                ; override;
    procedure OnAfterUpdateRecord        ; override;
    procedure OnArgument ( S: String )   ; override;
  private
    CodeParam 	: string;
    DateParam	: TDateTime;
    PeutFermer  : Boolean;
    FCanClose   : TCloseQueryEvent;

    procedure BFermeMouseDown(Sender: TObject; Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
    procedure ECanClose  (Sender : TObject; var CanClose : Boolean);
  end ;

Implementation

{---------------------------------------------------------------------------------------}
procedure TOM_COTATIONTAUX.OnNewRecord ;
{---------------------------------------------------------------------------------------}
begin
  Inherited ;
  SetField('TTA_CODE', CodeParam);
  SetField('TTA_DATE', DateParam);
  PeutFermer := False;
end ;

{---------------------------------------------------------------------------------------}
procedure TOM_COTATIONTAUX.OnAfterUpdateRecord;
{---------------------------------------------------------------------------------------}
begin
  Inherited ;
  SetControlEnabled('TTA_CODE', True);
  CodeParam := GetField('TTA_CODE');
  DateParam := GetField('TTA_DATE') + 1;
end;

{---------------------------------------------------------------------------------------}
procedure TOM_COTATIONTAUX.OnArgument ( S: String ) ;
{---------------------------------------------------------------------------------------}
var
  Str : string ;
begin
  Inherited ;
  Ecran.HelpContext := 150;
  TToolBarButton97(GetControl('BFERME')).OnMouseDown := BFermeMouseDown;
  Str := Trim(ReadtokenSt(S)) ;
  CodeParam := Trim(ReadtokenSt(S));

  Str := Trim(ReadtokenSt(S));
  if Str = '' then DateParam := Date
              else DateParam := StrToDate(Str);

  {Pour simuler une création en série : on ne peut fermer la fiche par le bouton Valider que
   si on n'est pas passer par le mode création}
  FCanClose := Ecran.OnCloseQuery;
  Ecran.OnCloseQuery := ECanClose;
  PeutFermer := True;
end ;

{---------------------------------------------------------------------------------------}
procedure TOM_COTATIONTAUX.ECanClose(Sender : TObject; var CanClose : Boolean);
{---------------------------------------------------------------------------------------}
begin
  CanClose := PeutFermer;
  if CanClose then FCanClose(Sender, Canclose);
end;

{---------------------------------------------------------------------------------------}
procedure TOM_COTATIONTAUX.BFermeMouseDown(Sender: TObject; Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
{---------------------------------------------------------------------------------------}
begin
  PeutFermer := True;
end;

Initialization
  registerclasses ( [ TOM_COTATIONTAUX ] ) ;
end.

