unit SelectDocSuggestion;

interface
uses SysUtils, Classes, UTOB, UTOF,Hctrls, HStatus, FactUtil, StdCtrls, math,
     EntGC, Ent1,AglInit,HTB97,grids,
{$IFDEF EAGLCLIENT}
     MaineAGL,
{$ELSE}
     db,{$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}FE_Main,
{$IFDEF V530}
      EdtEtat,EdtDoc,
{$ELSE}
      EdtREtat, EdtRDoc,
{$ENDIF}
{$ENDIF}
     HEnt1,Vierge;
TYPE
  TOF_GCSELDOCREA_MUL = Class (TOF)
  private
    TF : TFVierge;
    GS : THGrid;
    lesColonnes : string;
    BValider : TToolbarButton97;
    BSelectAll : TToolbarButton97;
    procedure definiLignescolonnesGS;
    procedure AffichageGS;
  public
    procedure OnArgument(stArgument : String); override;
    procedure OnLoad ; override ;
    procedure BValiderClick (Sender : Tobject);
    procedure BSelectAllClick (Sender : Tobject);
  end ;

function SelectionDocuments (LesNaturesAPrendre : string;DateDebut,DateFin: string; TOBListeDocuments : TOB) : boolean;

implementation
uses UTOF_VideInside;

function SelectionDocuments (LesNaturesAPrendre : string;DateDebut,DateFin: string; TOBListeDocuments : TOB) : boolean;
begin
  result := false;
  TRY
    TheToB := TOBListeDocuments;
    AGLLanceFiche('GC','GCSELDOCREA_MUL','','','ACTION=MODIFICATION');
    if TheTob <> nil then
    begin
      TOBListeDocuments.clearDetail;
      TOBListeDocuments.Dupliquer (TheTob,true,true);
      TheTOB.free;
      result := true;
    end;
  FINALLY
    TheTob := nil;
  END;
end;

{ TOF_GCSELDOCREA_MUL }

procedure TOF_GCSELDOCREA_MUL.definiLignescolonnesGS;
var i:integer;
    Nam,st : string;
    FF : string;
    NbColonnes : integer;
begin
GS.AllSelected := false;
st := LesColonnes;
NbColonnes := 0;
repeat
   Nam:=ReadTokenSt(St) ;
   if nam = '' then break;
   inc(NbCOlonnes);
until nam = '';
GS.VidePile (false);
GS.ColCount := NbColonnes;

if LaTob.Detail.count < 11 then GS.rowcount := 11
                           else GS.rowCount := LaTob.Detail.count+1;
FF:='#';
if V_PGI.OkDecV>0 then
   begin
    FF:='# ##0.';
    for i:=1 to V_PGI.OkDecV-1 do
       begin
       FF:=FF+'0';
       end;
    FF:=FF+'0';
   end;
st := LesColonnes;
i := 0;
repeat
   Nam:=ReadTokenSt(St) ;
   if nam = '' then break;
   if Nam='SELECT' then
      BEGIN
      GS.Cells[i,0]:=' ' ;
//      GS.ColFormats[i]:=-1;
      GS.ColWidths [i]:=10;
      GS.ColLengths [i] := -1;
      END;
   if Nam='GP_AFFAIRE' then
      begin
      GS.Cells[i,0]:='Affaire';
      GS.ColWidths [i]:=110;
      end;
   if Nam='GP_NATUREPIECEG' then
      begin
      GS.Cells[i,0]:='Nature';
      GS.ColFormats[i]:='CB=GCNATUREPIECEG||' ;
      GS.ColAligns [i] := taLeftJustify;
      GS.ColWidths [i]:=101;
      end;
   if Nam='GP_NUMERO' then
      begin
      GS.Cells[i,0]:='Numéro';
      GS.ColTypes[i]:='R' ;
      GS.ColAligns [i] := taRightJustify;
      GS.ColFormats[i]:='#0' ;
      GS.ColWidths [i]:=63;
      end;
   if Nam='GP_DEVISE' then
      begin
      GS.Cells[i,0]:='Devise';
      GS.ColFormats[i]:='CB=FEMDEVISES||' ;
      GS.ColAligns [i] := taCenter;
      GS.ColWidths [i]:=40;
      end;
   if Nam='GP_VIVANTE' then
      begin
      GS.Cells[i,0]:='Modifiable';
      GS.ColTypes[i]:='B' ;
      GS.colformats[i]:= inttostr(Integer(csCoche));
      GS.ColAligns [i] := taCenter;
      GS.ColWidths [i]:=19;
      end;
   if Nam='GP_DATEPIECE' then
      begin
      GS.Cells[i,0]:='Date';
      GS.ColTypes[i]:='D' ;
      GS.ColWidths [i]:=50;
      end;
   if Nam='GP_TIERS' then
      begin
      GS.Cells[i,0]:='Tiers';
      GS.ColWidths [i]:=63;
      end;
   if Nam='T_LIBELLE' then
      begin
      GS.Cells[i,0]:='Raison sociale';
      GS.ColWidths [i]:=151;
      end;
   if (Nam='GP_TOTALHT')  then
      BEGIN
      GS.Cells[i,0]:='Total HT';
      GS.ColAligns [i] := taRightJustify;
      GS.ColTypes[i]:='R';
      GS.ColFormats[i]:=FF ;
      GS.ColWidths [i]:=101;
      END;
   inc (i);
until nam = '';
end;


procedure TOF_GCSELDOCREA_MUL.BValiderClick(Sender: Tobject);
var Indice : integer;
    TobRenvoie,MaTob : TOB;
begin
  TOBRenvoie := TOB.Create ('LE RETOUR',nil,-1);
  for indice := 1 to GS.RowCount do
  begin
//    GS.GotoLeBookmark(indice);
    if (GS.IsSelected(indice)) and (TOB(GS.Objects[0,Indice]) <> nil) then
    begin
      MaTob := TOB.Create ('PIECE',TOBRenvoie,-1);
      MaTOb.dupliquer (TOB(GS.Objects[0,Indice]),false,true);
    end;
  end;
  if TOBRenvoie.detail.count = 0 then TOBRenvoie.free
                                 else TheTOb := TOBRenvoie;
  TFVierge(Ecran).close;
end;

procedure TOF_GCSELDOCREA_MUL.BSelectAllClick(Sender: Tobject);
begin
  GS.ClearSelected;
  GS.AllSelected := BSelectAll.Down;
end;

procedure TOF_GCSELDOCREA_MUL.OnArgument(stArgument: String);
begin
  inherited;
  LesColonnes := 'SELECT;GP_AFFAIRE;GP_NATUREPIECEG;GP_NUMERO;GP_DEVISE;GP_VIVANTE;GP_DATEPIECE;GP_TIERS;T_LIBELLE;GP_TOTALHT' ;
  TF := TFVierge(Ecran);
  GS := THGrid(getControl('GS'));
  BValider := TToolBarButton97(Getcontrol('Bouvrir'));
  BValider.OnClick := BValiderClick;
  BSelectAll := TToolBarButton97(Getcontrol('BSelectAll'));
  BSelectAll.OnClick := BSelectAllClick;
  definiLignescolonnesGS;
  TToolBarButton97(Getcontrol('BVALIDER')).visible := false;
end;

procedure TOF_GCSELDOCREA_MUL.OnLoad;
begin
  inherited;
  AffichageGS;
end;

procedure TOF_GCSELDOCREA_MUL.AffichageGS;
var indice : integer;
begin
if LATOB.detail.count = 0 then exit;
for Indice := 0 to LATOb.detail.count -1 do
    begin
    LATOb.Detail[Indice].PutLigneGrid (GS,Indice+1,false,false,LesColonnes);
    GS.Objects[0,Indice+1] := LATOB.detail[Indice];
    end;
// resize de la grid
TF.HMTrad.ResizeGridColumns (GS);
end;

Initialization
Registerclasses([TOF_GCSELDOCREA_MUL]) ;

end.
