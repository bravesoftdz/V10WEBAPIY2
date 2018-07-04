{***********UNITE*************************************************
Auteur  ...... : 
Créé le ...... : 18/04/2018
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : BTPARAMPHASES ()
Mots clefs ... : TOF;BTPARAMPHASES
*****************************************************************}
Unit BTPARAMPHASES_TOF ;

Interface

Uses StdCtrls, 
     Controls, 
     Classes, 
{$IFNDEF EAGLCLIENT}
     db, 
     uDbxDataSet, 
     mul, 
{$else}
     eMul, 
{$ENDIF}
     uTob, 
     forms,
     sysutils, 
     ComCtrls,
     HCtrls, 
     HEnt1, 
     HMsgBox,
     Graphics,
     Grids,
     Types,
     Vierge, 
     UTOF ; 

const
  TheChamps = ' ;REFERENCE;LIBELLE;VISIBLE;SAISISSABLE;';

Type
  TOF_BTPARAMPHASES = Class (TOF)
    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnDisplay                ; override ;
    procedure OnClose                  ; override ;
    procedure OnCancel                 ; override ;
  private
    GS : THGrid;
    procedure Definigrid;
    procedure AfficheLagrid;
    procedure SetEvents;
    procedure GSDblClick (Sender :TObject);
    procedure GSPostDrawCell(ACol, ARow: Longint; Canvas: TCanvas; AState: TGridDrawState);
    procedure SettrueFalse(TOBL : TOB;Arow,Acol : integer);
    procedure AppliqueVisibleAuxFilles( Niveau,Arow: Integer ; Statut : boolean);
    procedure AppliqueSaisieAuxFilles( Niveau,Arow : Integer; Statut : boolean);

  end ;

Implementation

procedure TOF_BTPARAMPHASES.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_BTPARAMPHASES.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_BTPARAMPHASES.OnUpdate ;
begin
  Inherited ;
  LaTOB.SetString('MODIF','X');
end ;

procedure TOF_BTPARAMPHASES.OnLoad ;
begin
  Inherited ;
end ;

procedure TOF_BTPARAMPHASES.OnArgument (S : String ) ;
begin
  Inherited ;
  GS := THGrid(GetControl('GS'));
  Definigrid;
  AfficheLagrid;
  SetEvents;
end ;

procedure TOF_BTPARAMPHASES.OnClose ;
begin
  Inherited ;
end ;

procedure TOF_BTPARAMPHASES.OnDisplay () ;
begin
  Inherited ;
end ;

procedure TOF_BTPARAMPHASES.OnCancel () ;
begin
  Inherited ;
end ;

procedure TOF_BTPARAMPHASES.Definigrid;
begin
  GS.VidePile(false);
  //
  GS.ColCount := 5;
  GS.Cells[0,0] := '';
  GS.ColWidths[0] := 20;
  GS.ColAligns[0] := taRightJustify;
  //
  GS.Cells[1,0] := 'Référence';
  GS.ColWidths[1] := 30;
  GS.ColLengths[1] := 18;
  GS.ColAligns[1] := taLeftJustify;
  //
  GS.Cells[2,0] := 'Libellé';
  GS.ColWidths[2] := 70;
  GS.ColLengths[2] := 70;
  GS.ColAligns[2] := taLeftJustify;
  //
  GS.Cells[3,0] := 'Visible';
  GS.ColWidths[3] := 30;
  GS.ColAligns[3] := taCenter;
  GS.ColTypes [3] := 'B';
  GS.ColFormats[3]:=IntToStr(Integer(csCoche));
  //
  GS.Cells[4,0] := 'Saisissable';
  GS.ColWidths[4] := 30;
  GS.ColTypes [4] := 'B';
  GS.ColAligns[4] := taCenter;
  GS.ColFormats[4]:=IntToStr(Integer(csCoche));

end;

procedure TOF_BTPARAMPHASES.AfficheLagrid;
var  II : Integer;
begin
  GS.RowCount := LaTOB.Detail.Count+1;
  for II := 0 to LaTOB.Detail.Count -1 do
  begin
    LaTOB.detail[II].PutLigneGrid(GS,II+1,false,false,TheChamps); 
  end;
  TFVierge (Ecran).HMTrad.ResizeGridColumns(GS); 
end;

procedure TOF_BTPARAMPHASES.SetEvents;
begin
  GS.OnDblClick := GSDblClick;
  GS.PostDrawCell := GSPostDrawCell;   
end;

procedure TOF_BTPARAMPHASES.GSDblClick (Sender :TObject);
var TOBL : TOB;
begin
  TOBL := LaTOB.Detail[GS.row-1];
  SettrueFalse(TOBL,GS.row,GS.col);
  AfficheLagrid;
end;

procedure TOF_BTPARAMPHASES.GSPostDrawCell(ACol, ARow: Integer; Canvas: TCanvas; AState: TGridDrawState);
var ARect : Trect;
    decalage : integer;
    TOBL : TOB;
    TheText : string;
begin
  if (Acol = 0) or (Arow = 0) then Exit;
  ARect := GS.CellRect(ACol, ARow);
  if Acol = 2 then
  begin
    TOBL := LaTOB.detail[Arow-1];
    Decalage := 0;
    if TOBL.GetInteger('NIVEAU') >0 then
    begin
      Decalage := (TOBL.GetInteger('NIVEAU')-1) * 3 * canvas.TextWidth('w');
    end;
    TheText := TOBL.getvalue('LIBELLE');
    Canvas.FillRect(ARect);
    GS.Canvas.Brush.Style := bsSolid;
    GS.Canvas.TextOut (Arect.left + Decalage+1,Arect.Top +2 ,Thetext);
  end;

end;

procedure TOF_BTPARAMPHASES.SettrueFalse(TOBL: TOB; Arow,Acol: integer);
var Statut : boolean;
    Niveau : integer;
begin
  if Acol = 3 then // VISIBLE
  begin
    Niveau := TOBL.GetInteger('NIVEAU');
    if TOBL.GetBoolean('VISIBLE') then
    begin
      Statut := false;
      // Passage en invisible et non saisissable
    end else
    begin
      Statut := true;
    end;
    TOBL.SetBoolean('VISIBLE',Statut);
    TOBL.SetBoolean('SAISISSABLE',Statut);
    AppliqueVisibleAuxFilles( Niveau,Arow,Statut);
  end else if Acol = 4 then // SAISISSABLE
  begin
    Niveau := TOBL.GetInteger('NIVEAU');
    if TOBL.GetBoolean('SAISISSABLE') then
    begin
      Statut := false;
      // Passage en invisible et non saisissable
    end else
    begin
      Statut := true;
    end;
    TOBL.SetBoolean('SAISISSABLE',Statut);
    AppliqueSaisieAuxFilles( Niveau,Arow,Statut);
  end;
end;


procedure TOF_BTPARAMPHASES.AppliqueSaisieAuxFilles(Niveau, Arow: Integer; Statut: boolean);
var II : Integer;
    TheC : TOB;
begin
  if Arow >= LaTOB.Detail.count Then Exit;
  for II := Arow to LaTOB.Detail.count -1 do
  begin
    TheC := LaTOB.detail[II];
    if TheC.GetInteger('NIVEAU')=Niveau then break;
    TheC.SetBoolean('SAISISSABLE',Statut);
  end;
end;

procedure TOF_BTPARAMPHASES.AppliqueVisibleAuxFilles(Niveau, Arow: Integer; Statut: boolean);
var II : Integer;
    TheC : TOB;
begin
  if Arow >= LaTOB.Detail.count Then Exit;
  for II := Arow to LaTOB.Detail.count -1 do
  begin
    TheC := LaTOB.detail[II];
    if TheC.GetInteger('NIVEAU')=Niveau then break;
    TheC.SetBoolean('VISIBLE',Statut);
    TheC.SetBoolean('SAISISSABLE',Statut);
  end;
end;

Initialization
  registerclasses ( [ TOF_BTPARAMPHASES ] ) ;
end.
