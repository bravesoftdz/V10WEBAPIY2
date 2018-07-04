unit galTomUserConf;
// TOM de la table USERCONF
// #### source OBSOLETE, ne plus utiliser !!

interface
uses  Windows,StdCtrls,Controls,Classes,forms,sysutils,ComCtrls,
      HCtrls,HEnt1,HMsgBox,UTOM, Dialogs,HDB, graphics,
{$IFNDEF EAGLCLIENT}
      dbgrids,
{$ENDIF}
      grids;

Type
  TOM_UserConf = Class (TOM)
    procedure OnArgument(Arguments : String ) ; override ;
{$IFDEF EAGLCLIENT}
    procedure DessineCell(ACol, ARow: Longint; Canvas: TCanvas; AState: TGridDrawState);
{$ELSE}
    procedure DessineCell(Sender: TObject;const Rect: TRect; DataCol: Integer; Column: TColumn;
                           State: TGridDrawState);
{$ENDIF}
  private
    Liste : THDBGrid; // en eagl, casté en THGrid par HDb
  END;

////////////// IMPLEMENTATION //////////////
implementation

procedure TOM_UserConf.OnArgument(Arguments: String);
begin
inherited;
  Liste:=THDBGrid(GetControl('FListe')) ;
{$IFDEF EAGLCLIENT}
  Liste.PostDrawCell := DessineCell;
{$ELSE}
  Liste.OnDrawColumnCell := DessineCell;
{$ENDIF}
end;

{$IFDEF EAGLCLIENT}
procedure TOM_UserConf.DessineCell(ACol, ARow: Longint; Canvas: TCanvas; AState: TGridDrawState);
var
  TT,Texte : string;
  Rect: TRect;
begin
  // #### A FAIRE eAGL ####
  // dans Titres, il y a les champs issus de la liste
  if Liste.Titres[ACol]='UCO_USER' then TT := 'TTUTILISATEUR'
  else if Liste.Titres[ACol]='UCO_GROUPECONF' then TT := 'GROUPECONF';
  Texte := RechDom(TT,Liste.Cells[ACol, ARow],FALSE);
  Canvas.FillRect(Canvas.ClipRect);
  Canvas.TextOut( Canvas.ClipRect.Left+3,Canvas.ClipRect.Top+3,Texte);
end;
{$ELSE}
procedure TOM_UserConf.DessineCell(Sender: TObject;
  const Rect: TRect; DataCol: Integer; Column: TColumn;
  State: TGridDrawState);
var
  TT,Texte : string;
begin
  if Column.Field.DisplayName='UCO_USER' then TT := 'TTUTILISATEUR'
  else if Column.Field.DisplayName='UCO_GROUPECONF' then TT := 'GROUPECONF';
  Texte := RechDom(TT,Column.Field.AsString,FALSE);
  Liste.Canvas.FillRect(Rect);
  Liste.Canvas.TextOut( Rect.Left+3,Rect.Top+3,Texte);
end;
{$ENDIF}

Initialization
registerclasses([TOM_UserConf]) ;
end.
