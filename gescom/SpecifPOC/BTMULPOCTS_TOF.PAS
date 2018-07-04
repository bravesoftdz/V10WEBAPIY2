{***********UNITE*************************************************
Auteur  ...... :
Créé le ...... : 27/12/2017
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : BTMULPOCTS ()
Mots clefs ... : TOF;BTMULPOCTS
*****************************************************************}
Unit BTMULPOCTS_TOF ;

Interface

Uses StdCtrls,
     Controls,
     Classes,
{$IFNDEF EAGLCLIENT}
     db,
     uDbxDataSet,
     mul,
     fe_main,
{$else}
     eMul,
{$ENDIF}
     AGLInit,
     uTob,
     forms,
     sysutils,
     ComCtrls,
     HCtrls,
     HEnt1,
     HMsgBox,
     Vierge,
     HTB97,
     Hpanel,
     UTOF ;

Type
  TOF_BTMULPOCTS = Class (TOF)
    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnDisplay                ; override ;
    procedure OnClose                  ; override ;
    procedure OnCancel                 ; override ;
  private
    TOBPiece : TOB;
    TOBDet : TOB;
    GS : THGrid;
    procedure PrepareLaGrille;
    procedure AfficheLaGrille;
    procedure BINSERTCLick (SEnder : TObject);
    procedure BDELETECLick (Sender : TObject);
    procedure GSDblClick (Sender : Tobject);
    procedure CalcMontantCumul (TOBD : TOB);
    procedure CalcMontantGlob;

  end ;

const
  TheChamps = 'RET;BLE_REFERENCETS;BLE_DATE;BLE_MONTANT;';
  
Implementation

procedure TOF_BTMULPOCTS.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_BTMULPOCTS.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_BTMULPOCTS.OnUpdate ;
begin
  Inherited ;
end ;

procedure TOF_BTMULPOCTS.OnLoad ;
begin
  Inherited ;
end ;

procedure TOF_BTMULPOCTS.OnArgument (S : String ) ;
begin
  Inherited ;
  TOBDet := LaTOB;
  TOBPiece := TOB(TOBDet.data);
  THpanel(GetControl('PDESC')).Caption := LaTOB.getString('DESIGNATION');
  GS := THGrid(GetControl('GS'));
  TToolbarButton97 (GetControl('Binsert')).OnClick := BINSERTCLick;
  TToolbarButton97 (GetControl('BDelete')).OnClick := BDELETECLick;
  PrepareLaGrille;
  AfficheLaGrille;
  GS.OnDblClick := GSDblClick;
end ;

procedure TOF_BTMULPOCTS.OnClose ;
begin
  Inherited ;
end ;

procedure TOF_BTMULPOCTS.OnDisplay () ;
begin
  Inherited ;
end ;

procedure TOF_BTMULPOCTS.OnCancel () ;
begin
  Inherited ;
end ;

procedure TOF_BTMULPOCTS.AfficheLaGrille;
var NbLig : Integer;
    II : Integer;
begin
  NbLig := TOBDet.detail.Count; if NbLig = 0 then NbLig := 1;
  GS.RowCount := NbLig +1;
  //
  if TOBDet.Detail.count > 0 then
  begin
    for II := 0 to TOBDet.detail.Count -1 do
    begin
      TOBDet.detail[II].PutLigneGrid(GS,II+1,False,False,TheChamps);
    end;
  end;
  TFVierge(Ecran).HMTrad.ResizeGridColumns(GS);
end;

procedure TOF_BTMULPOCTS.BDELETECLick(Sender: TObject);
var OneTOB : TOB;
begin
  if PGIAsk('Etes-vous sur de vouloir supprimer cet Avenant/TS ?') <> mryes then exit;
  OneTOB := TOBDet.Detail[GS.row-1];
  OneTOB.free;
  GS.DeleteRow(GS.Row -1);
  AfficheLaGrille;
  CalcMontantGlob;
  //
end;

procedure TOF_BTMULPOCTS.BINSERTCLick(SEnder: TObject);
var OneTOB,OneDetail : TOB;
begin
  OneTOB := TOB.Create ('BLIGNEETS',nil,-1); OneTOB.AddChampSupValeur ('OK','-');
  OneTOB.SetInteger('BLE_NUMORDRE',TOBDet.GetInteger('NUMORDRE'));
  TheTOB := OneTOB;
  OneTOB.Data := TOBDet;
  TOBDet.Data := TOBPiece;
  AGLLanceFiche('BTP','BTSAISPOCTS','','','ACTION=CREATION');
  TheTOB := nil;
  OneTOB.Data := nil;
  TOBDet.Data := nil;
  if OneTOB.GetString('OK') = 'X' then
  begin
    CalcMontantCumul (OneTOB);
    OneTOB.ChangeParent(TOBDet,-1);
    AfficheLaGrille;
    CalcMontantGlob;
  end else OneTOB.Free;

end;

procedure TOF_BTMULPOCTS.GSDblClick(Sender: Tobject);
var OneTOB : TOB;
begin
  OneTOB := TOBDet.Detail[GS.row-1];
  OneTOB.Data := TOBDet;
  TOBDet.Data := TOBPiece;
  if not OneTOB.FieldExists('OK') then OneTOB.AddChampSupValeur ('OK','-')
                                  else OneTOB.SetString ('OK','-');
  TheTOB := OneTOB;
  AGLLanceFiche('BTP','BTSAISPOCTS','','','ACTION=MODIFICATION');
  TheTOB := nil;
  OneTOB.Data := nil;
  if OneTOB.GetString('OK')= 'X' then
  begin
    CalcMontantCumul (OneTOB);
    AfficheLaGrille;
    CalcMontantGlob;
  end;
end;

procedure TOF_BTMULPOCTS.CalcMontantCumul(TOBD: TOB);
var II : Integer;
begin
  TOBD.SetDouble('BLE_MONTANT',0);
  for II := 0 to TOBD.detail.count -1 do
  begin
    TOBD.setDouble('BLE_MONTANT',TOBD.GetDouble('BLE_MONTANT')+TOBD.detail[II].GetDouble('BLT_MONTANTHTDEV'));
  end;
end;

procedure TOF_BTMULPOCTS.CalcMontantGlob;
var II : Integer;
begin
  TOBDet.setDouble('SUMTOTALTS',0);
  for II := 0 to TOBDet.Detail.count -1 do
  begin
    TOBDet.setDouble('SUMTOTALTS',TOBDet.GetDouble('SUMTOTALTS')+TOBDet.detail[II].getDouble('BLE_MONTANT'));
  end;
end;

procedure TOF_BTMULPOCTS.PrepareLaGrille;
var NbLig : Integer;
    II : Integer;
    FFP : string;
begin
  FFP := '#';
  if V_PGI.OkDecV > 0 then
  begin
    FFP := '0.';
    for II := 1 to V_PGI.OkDecV - 1 do
    begin
      FFP := FFP + '#';
    end;
    FFP := FFP + '0';
  end;

  GS.VidePile(false);
  GS.ColCount := 4;
  NbLig := TOBDet.detail.Count; if NbLig = 0 then NbLig := 1;
  GS.RowCount := NbLig +1;
  GS.Cells[0,0] := '';
  GS.ColWidths[0] := 20;
  //
  GS.Cells[1,0] := 'Reference';
  GS.ColWidths[1] := 50;
  GS.ColAligns  [1] := taLeftJustify;
  //
  GS.Cells[2,0] := 'Date';
  GS.ColWidths[2] := 30;
  GS.ColTypes[2] := 'D';
  GS.ColFormats[2] := ShortdateFormat;
  GS.ColAligns  [2] := taLeftJustify;
  //
  GS.Cells[3,0] := 'Montant';
  GS.ColWidths[3] := 30;
  GS.ColFormats[3] := FFP+';'+FFP+'; ;';
  GS.ColAligns  [3] := taRightJustify;
  //
end;

Initialization
  registerclasses ( [ TOF_BTMULPOCTS ] ) ;
end.
