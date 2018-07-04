{***********UNITE*************************************************
Auteur  ...... : Vincent Laroche
Créé le ...... : 20/02/2003
Modifié le ... : 12/08/2004
Description .. : Passage en eAGL
Suite ........ : CA - 12/08/2004
Suite ........ : -> FQ 13071 : lien avec GENERAUXREF en contexte PCL
Mots clefs ... : 
*****************************************************************}
unit ImpPrefG;

interface

uses
  Windows,
  Messages,
  SysUtils,
  Classes,
  Graphics,
  Controls,
  Forms,
  Dialogs,
{$IFDEF EAGLCLIENT}
  UTOB,
{$ELSE}
  {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF}
{$ENDIF}
  StdCtrls,
  Buttons,
  ComCtrls,
  HRichEdt,
  Hctrls,
  Mask,
  ExtCtrls,
  HSysMenu,
  HRichOLE,
  MajTable,
  Hqry,
  HEnt1;

Function FicheImportPlanRef(NumPlan : integer ; Compte : String) : Integer ;

type
  TFImporPRef = class(TForm)
    Pappli: TPanel;
    HGBOptionaxes: TGroupBox;
    PR_VENTILABLE1: TCheckBox;
    PR_VENTILABLE2: TCheckBox;
    PR_VENTILABLE3: TCheckBox;
    PR_VENTILABLE4: TCheckBox;
    PR_VENTILABLE5: TCheckBox;
    HGBOptionImpression: TGroupBox;
    PR_SOLDEPROGRESSIF: TCheckBox;
    PR_TOTAUXMENSUELS: TCheckBox;
    PR_SAUTPAGE: TCheckBox;
    PR_CENTRALISABLE: TCheckBox;
    TG_BLOCNOTE: TGroupBox;
    PR_BLOCNOTE: THRichEditOLE;
    GBDescrip: TGroupBox;
    TPR_COMPTE: THLabel;
    PR_COMPTE: TEdit;
    tPR_ABREGE: THLabel;
    PR_ABREGE: TEdit;
    PR_LIBELLE: TEdit;
    TPR_LIBELLE: THLabel;
    TPR_SENS: THLabel;
    TPR_NATUREGENE: THLabel;
    GBCarac: TGroupBox;
    PR_COLLECTIF: TCheckBox;
    PR_POINTABLE: TCheckBox;
    PR_LETTRABLE: TCheckBox;
    PBouton: TPanel;
    BImprimer: TBitBtn;
    BValider: THBitBtn;
    HelpBtn: THBitBtn;
    BFerme: THBitBtn;
    PR_SENS: THValComboBox;
    PR_NATUREGENE: THValComboBox;
    CbMess: TComboBox;
    Label1: TLabel;
    HMTrad: THSystemMenu;
    procedure FormShow(Sender: TObject);
    procedure HelpBtnClick(Sender: TObject);
  private
    NumPlan : Integer ;
    Compte : String ;
    { Déclarations privées }
  public
    { Déclarations publiques }
  end;


implementation

{$R *.DFM}

Function FicheImportPlanRef(NumPlan : Integer ; Compte : String) : Integer;
var FImporPRef : TFImporPRef ;
BEGIN
Result:=0 ;
FImporPRef:=TFImporPRef.Create(Application) ;
  Try
   FImporPRef.NumPlan:=NumPlan ;
   FImporPRef.Compte:=Compte ;
   FImporPRef.BImprimer.Visible:=False ;
   if FImporPRef.ShowModal=mrOk then Result:=1 ;
  Finally
   FImporPRef.Free ;
  End ;
END ;


procedure TFImporPRef.FormShow(Sender: TObject);
var
  Q : TQuery;
begin
  if ctxPCL in V_PGI.PGIContexte then
    Q := OpenSQL('SELECT GER_NUMPLAN,GER_GENERAL,GER_LIBELLE,GER_ABREGE,GER_NATUREGENE,GER_CENTRALISABLE,'+
      'GER_SOLDEPROGRESSIF,GER_SAUTPAGE,GER_TOTAUXMENSUELS,GER_COLLECTIF,"",GER_SENS,GER_LETTRABLE,'+
      'GER_POINTABLE,GER_VENTILABLE1,GER_VENTILABLE2,GER_VENTILABLE3,GER_VENTILABLE4,GER_VENTILABLE5 '+
      'FROM GENERAUXREF WHERE GER_NUMPLAN='+IntToStr(NumPlan)+' AND GER_GENERAL="'+Compte+'"',True)
  else
  { FQ 20958 BVE 16.07.07 }
    Q := OpenSQL('SELECT * FROM PLANREF WHERE PR_NUMPLAN='+IntToStr(NumPlan)+' AND PR_COMPTE="'+Compte+'" AND PR_PREDEFINI="STD"',True);
  if Q.EOF then
  begin
    Ferme(Q);
    Q := OpenSQL('SELECT * FROM PLANREF WHERE PR_NUMPLAN='+IntToStr(NumPlan)+' AND PR_COMPTE="'+Compte+'" AND PR_PREDEFINI="CEG"',True);
  end;
  { END FQ 20958 }
  if (not Q.EOF) then begin
    PR_COMPTE.Text := Q.Fields[1].AsString;
    PR_LIBELLE.Text := Q.Fields[2].AsString;
    PR_ABREGE.Text := Q.Fields[3].AsString;
    { FQ 20958 BVE 16.07.07
    PR_NATUREGENE.Text := Q.Fields[4].AsString; }
    PR_NATUREGENE.Value := Trim(Q.Fields[4].AsString);     
    { END FQ 20958 }
    PR_CENTRALISABLE.Checked := (Q.Fields[5].AsString = 'X');
    PR_SOLDEPROGRESSIF.Checked := (Q.Fields[6].AsString = 'X');
    PR_SAUTPAGE.Checked := (Q.Fields[7].AsString = 'X');
    PR_TOTAUXMENSUELS.Checked := (Q.Fields[8].AsString = 'X');
    PR_COLLECTIF.Checked := (Q.Fields[9].AsString = 'X');
    PR_BLOCNOTE.Text := Q.Fields[10].AsString;  
    { FQ 20958 BVE 16.07.07
    PR_SENS.Text := Q.Fields[11].AsString;  }
    PR_SENS.Value := Trim(Q.Fields[11].AsString);
    { END FQ 20958 }      
    PR_LETTRABLE.Checked := (Q.Fields[12].AsString = 'X');
    PR_POINTABLE.Checked := (Q.Fields[13].AsString = 'X');
    PR_VENTILABLE1.Checked := (Q.Fields[14].AsString = 'X');
    PR_VENTILABLE2.Checked := (Q.Fields[15].AsString = 'X');
    PR_VENTILABLE3.Checked := (Q.Fields[16].AsString = 'X');
    PR_VENTILABLE4.Checked := (Q.Fields[17].AsString = 'X');
    PR_VENTILABLE5.Checked := (Q.Fields[18].AsString = 'X');
  end;
  Ferme(Q);
  Pappli.Enabled:=False ;
  Caption:=CbMess.Items[0]+' '+Compte+' '+CbMess.Items[1]+IntToStr(NumPlan) ;
  {$IFDEF CCS3}
    PR_VENTILABLE2.Visible:=False ; PR_VENTILABLE3.Visible:=False ;
    PR_VENTILABLE4.Visible:=False ; PR_VENTILABLE5.Visible:=False ;
  {$ENDIF}
  UpdateCaption(Self) ;
end;

procedure TFImporPRef.HelpBtnClick(Sender: TObject);
begin
CallHelpTopic(Self) ;
end;

end.
