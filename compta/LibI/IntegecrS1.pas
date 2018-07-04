unit IntegecrS1;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  Grids, StdCtrls, HPanel, Buttons, HCtrls,ExtCtrls,HEnt1,DBTables, Db, Math,
  DBGrids, Spin, hmsgbox, MajTable, Hqry, DBCtrls,
  Mask,CptesAssS1,Outils, HTB97, HSpeller, ComCtrls,ParamDat, HSysMenu, UiUtil,
  Menus,Filtre,UTob,ImGenEcr,HStatus,ParamSoc,UtilPGI,UtInput,ecr_edit, ImEnt, imOutGen ;

type
  TIntegration = class(TForm)
    HM: THMsgBox;
    HMTrad: THSystemMenu;
    POPF: TPopupMenu;
    BCreerFiltre: TMenuItem;
    BSaveFiltre: TMenuItem;
    BDelFiltre: TMenuItem;
    BRenFiltre: TMenuItem;
    BNouvRech: TMenuItem;
    Dock971: TDock97;
    ToolWindow971: TToolWindow97;
    BValider: TToolbarButton97;
    BFerme: TToolbarButton97;
    HelpBtn: TToolbarButton97;
    BImprimer: TToolbarButton97;
    PanelEch: TPanel;
    HLabel6: THLabel;
    DateDebutEch: THCritMaskEdit;
    HLabel7: THLabel;
    DateFinEch: THCritMaskEdit;
    LibelleCBanque: THLabel;
    GEcheance: TGroupBox;
    tJEcheance: THLabel;
    tCTva: THLabel;
    cbEcrEcheance: TCheckBox;
    JEcheance: THValComboBox;
    CTva: THCritMaskEdit;
    GBanque: TGroupBox;
    tCBanque: THLabel;
    tJBanque: THLabel;
    cbEcrPaiement: TCheckBox;
    JBanque: THValComboBox;
    CBanque: THCritMaskEdit;
    PanelDot: TPanel;
    GroupBox2: TGroupBox;
    HLabelDateCalcul: THLabel;
    HLabel2: THLabel;
    DateCalcul: THCritMaskEdit;
    DateGeneration: THCritMaskEdit;
    GroupBox1: TGroupBox;
    HLabel4: THLabel;
    HLabel5: THLabel;
    LibelleUnique: TEdit;
    CodeJournal: THValComboBox;
    HLabel1: THLabel;
    cTypEcrEch: THValComboBox;
    HLabel3: THLabel;
    cTypEcrDot: THValComboBox;
    bZoom: TToolbarButton97;
    procedure FormShow(Sender: TObject);
    procedure cbEcrEcheanceClick(Sender: TObject);
    procedure cbEcrPaiementClick(Sender: TObject);
    procedure DateKeyPress(Sender: TObject; var Key: Char);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure BValiderClick(Sender: TObject);
    procedure JBanqueChange(Sender: TObject);
    procedure bZoomClick(Sender: TObject);
    procedure ToolWindow971Resize(Sender: TObject);
  private
    { Déclarations privées }
    fTypeEcr : TypeIntegration;
    fListeEcriture : TList;
    fListeImmo : TStrings;
    procedure PrepareEcrituresDot;
    procedure PrepareEcrituresEch;
    procedure IntegrationComptaPourTrans ;
    procedure IntegrationCompta(Visu: boolean);
    procedure MajTableEcheance;
    procedure MarqueIntegreEch(CodeImmo: string) ;
    function ControleZoneEcheance : boolean;
    function ControleZoneDotation: boolean;
    function GetCompteBanqueContrepartie: string;
  public
    { Déclarations publiques }
  end;

procedure IntegrationEcritures (Op : TypeIntegration;ListeImmo : TStrings;bCompta,bDetail : boolean);

implementation

uses ImpExp,PrintDBG,(*VerifEcr,*)ImRapInt,ImVisuEcr;

{$R *.DFM}
procedure IntegrationEcritures (Op : TypeIntegration;ListeImmo : TStrings;bCompta,bDetail : boolean);
var
  Integration: TIntegration;
  PP : THpanel;
begin
//  if GetParamSoc('SO_EXOCLOIMMO')=VHImmo^.Encours.Code then exit;
  Integration := TIntegration.Create(Application);
  Integration.fTypeEcr:=Op;
  Integration.fListeImmo:=ListeImmo;
//  Integration.rbDetail.Checked := bDetail;
//  Integration.fbCompta := bCompta;
  PP:=FindInsidePanel ;
  if PP=Nil then
  begin
    try
      Integration.ShowModal;
    finally
      Integration.Free;
    end;
  end else
  begin
    InitInside(Integration,PP) ;
    Integration.Show ;
  end;
end;

function CompareSuivantCompteRef (Item1,Item2:Pointer) : integer;
var Ecr1,Ecr2 : TLigneEcriture;
begin
  Ecr1 := Item1;Ecr2 := Item2;
  if Ecr1.CompteRef > Ecr2.CompteRef then Result := 1
  else if Ecr1.CompteRef < Ecr2.CompteRef then Result := -1
  else
  begin
    if Ecr1.Compte > Ecr2.Compte then Result := 1
    else if Ecr1.Compte < Ecr2.Compte then Result := - 1
    else Result := 0;
  end;
end;

function GetCollectif(CpteAux: string): string ;
var Q: TQuery ;
begin
  result:='' ;
  Q:=OpenSql('SELECT T_COLLECTIF FROM TIERS WHERE T_AUXILIAIRE="'+CpteAux+'"',false) ;
  if not Q.Eof then result:=Q.Fields[0].AsString ;
  ferme(Q) ;
end ;

procedure TIntegration.FormShow(Sender: TObject);
var Action : TCloseAction;
begin
  fListeEcriture:=TList.Create ;
  if fTypeEcr = toDotation then
  begin
    Caption:=HM.Mess[3];
    PanelEch.Width:=0 ;
    ClientWidth:=PanelDot.Width ;
    DateCalcul.Text    :=DateToStr(VHImmo^.Encours.Fin);
    DateGeneration.Text:=DateToStr(VHImmo^.Encours.Fin);
  end
  else
  begin
    Caption := HM.Mess[4];
    PanelDot.Width:=0 ;
    ClientWidth:=PanelEch.Width ;
    DateDebutEch.Text:=DateToStr(VHImmo^.Encours.Deb);
    DateFinEch.Text  :=DateToStr(VHImmo^.Encours.Fin);
  end;
  if GetParamSoc('SO_COLLIMMO')='' then PGIBox('Le collectif d''immobilisation n''existe pas,#13veuillez le renseigner dans les paramètres société',caption) ;
//  else MajFiche;
end;

procedure TIntegration.FormClose(Sender: TObject;
  var Action: TCloseAction);
begin
  VideListeEcritures(fListeEcriture);
  fListeEcriture.Free ;
  if isInside(Self) then
    Action:=caFree ;
end;

procedure TIntegration.PrepareEcrituresDot;
var i :integer; ParamEcr: TParamEcr;
begin
  Screen.Cursor := crHourglass;
  VideListeEcritures(fListeEcriture);
  ParamEcr := TParamEcr.Create;
  try
    ParamEcr.Journal   := CodeJournal.Value;
    ParamEcr.Libelle   := LibelleUnique.Text;
    ParamEcr.Date      := StrToDate (DateGeneration.Text);
    ParamEcr.DateCalcul:= StrToDate (DateCalcul.Text);
    ParamEcr.bDotation := True;
    if fListeImmo = nil then
    begin
      InitMove (1,'');
      CalculEcrituresDotation (fListeEcriture,'',ParamEcr,StrToInt(cTypEcrDot.Value));
      FiniMove;
    end
    else
    begin
      InitMove(fListeImmo.Count,'');
      for i:=0 to fListeImmo.Count-1 do
        CalculEcrituresDotation (fListeEcriture,fListeImmo[i],ParamEcr,StrToInt(cTypEcrDot.Value));
      FiniMove;
    end;
  finally
    ParamEcr.Free;
  end ;
  Screen.Cursor:=SyncrDefault ;
end;

procedure TIntegration.PrepareEcrituresEch;
var i :integer;
    ParamEch,ParamPai : TParamEcr;
    dtEchMin, dtEchMax : TDateTime;
begin
  Screen.Cursor := crHourglass;
  VideListeEcritures(fListeEcriture);
  try
    if cbEcrEcheance.Checked then
    begin
      ParamEch          :=TParamEcr.Create;
      ParamEch.Journal  :=JEcheance.Value;
      ParamEch.CompteRef:=CTva.Text;
      ParamEch.bDotation:=False;
    end ;
    if cbEcrPaiement.Checked then
    begin
      ParamPai          :=TParamEcr.Create;
      ParamPai.Journal  :=JBanque.Value;
      ParamPai.CompteRef:=CBanque.Text;
      ParamPai.bDotation:=False;
    end ;
    if (ParamPai=nil) and (ParamEch=nil) then exit ;
    dtEchMin := StrToDate(DateDebutEch.Text);
    dtEchMax := StrToDate(DateFinEch.Text);
    if fListeImmo = nil then
    begin
      InitMove (1,'');
      CalculEcrituresEcheances ( fListeEcriture,'',ParamEch,ParamPai,StrToInt(cTypEcrEch.Value),dtEchMin,dtEchMax);
      FiniMove;
    end
    else
    begin
      InitMove(fListeImmo.Count,'');
      for i:=0 to fListeImmo.Count - 1 do
        CalculEcrituresEcheances ( fListeEcriture,fListeImmo[i],ParamEch,ParamPai,StrToInt(cTypEcrEch.Value),dtEchMin,dtEchMax);
      FiniMove;
    end;
  finally
    if ParamEch<>nil then ParamEch.Free;
    if ParamPai<>nil then ParamPai.Free;
  end ;
  Screen.Cursor:=SyncrDefault ;
end;

procedure TIntegration.cbEcrEcheanceClick(Sender: TObject);
begin
  JEcheance.Enabled :=cbEcrEcheance.Checked;
  tJEcheance.Enabled:=cbEcrEcheance.Checked;
  CTva.Enabled      :=cbEcrEcheance.Checked;
  tCTva.Enabled     :=cbEcrEcheance.Checked;
end;
procedure TIntegration.cbEcrPaiementClick(Sender: TObject);
begin
  JBanque.Enabled :=cbEcrPaiement.Checked;
  tJBanque.Enabled:=cbEcrPaiement.Checked;
  tCBanque.Enabled:=cbEcrPaiement.Checked;
end;

procedure TIntegration.DateKeyPress(Sender: TObject;var Key: Char);
begin
  ParamDate (Self, Sender, Key);
end;

procedure TIntegration.BValiderClick(Sender: TObject);
var mr,i : integer;
    ARecord : TCRPiece;
begin
  if GetParamSoc('SO_COLLIMMO')='' then PGIBox('Le collectif d''immobilisation n''existe pas,#13veuillez le renseigner dans les paramètres société',caption) ;
  if (fTypeEcr=toDotation) then
  begin
    if not ControleZoneDotation then exit;
    PrepareEcrituresDot ;
  end
  else
  begin
    if not ControleZoneEcheance then exit;
    PrepareEcrituresEch ;
  end ;

  if (HM.Execute (8,Caption,'')=mrYes) then
  begin
    if Blocage(['nrCloture','nrBatchImmo'],True,'nrBatchImmo') then exit;
    if Transactions(IntegrationComptaPourTrans,5)<>oeOK then
    begin
      HM.Execute(9,Caption,'')
    end
    else
    begin
      HM.Execute(23,Caption,'') ;
      ModalResult:=mrOk;
    end;
    Bloqueur('nrBatchImmo',False) ;
  end ;
end;

procedure TIntegration.IntegrationComptaPourTrans ;
begin
  IntegrationCompta(false) ;
end ;

procedure TIntegration.IntegrationCompta(Visu: boolean);
var i, j : integer; ARecord : TLigneEcriture; //ARecordAna : TAna;
    ToutesLesTobs,E_s,Y_,Ecr_: tob; io : TIOS1 ; Okok,ArreterSiErreur: boolean ; tDate: TDateTime ; tJal: string ;
begin
  tDate:=iDate1900 ; tJal:='' ;
  ToutesLesTobs:=Tob.Create('UNKOWN',nil,-1) ;
  try
    for i := 0 to fListeEcriture.Count - 1 do
    begin
      ARecord := fListeEcriture.Items[i];
      if (tdate<>ARecord.Date) or (tJal<>ARecord.CodeJournal) then E_s:=Tob.Create('ECRITURE_S',ToutesLesTobs,-1) ;
      tdate:=ARecord.Date ; tJal:=ARecord.CodeJournal ;
      Ecr_:=Tob.create('ECRITURE',E_s,-1) ;
      if ARecord.Auxi<>'' then ARecord.Compte:=GetCollectif(ARecord.Auxi) ;
      Ecr_.PutValue('E_GENERAL',ARecord.Compte) ;
      Ecr_.PutValue('E_AUXILIAIRE',ARecord.Auxi) ;
      Ecr_.PutValue('E_NUMEROPIECE',-1) ;
      Ecr_.PutValue('E_LIBELLE',ARecord.Libelle) ;
      Ecr_.PutValue('E_DATECOMPTABLE',ARecord.Date) ;
      Ecr_.PutValue('E_JOURNAL',ARecord.CodeJournal) ;
      Ecr_.PutValue('E_DEBITDEV',ARecord.Debit) ;
      Ecr_.PutValue('E_CREDITDEV',ARecord.Credit) ;
      Ecr_.PutValue('E_TYPESAISIE',8) ;  //YCP 07/12/00 a voir type saisie => IMMO
      Ecr_.PutValue('E_DEVISE',V_PGI.DevisePivot) ;
      Ecr_.PutValue('E_ANA', '-');
(*      if ARecord.Axes[0]<>nil then
      begin
        Ecr_.PutValue('E_ANA', 'X');
        for j := 0 to ARecord.Axes[0].Count - 1 do
        begin
          ARecordAna := ARecord.Axes[0].Items[j];
          Y_:=Tob.Create('ANALYTIQ',Ecr_,-1) ;
          Y_.PutValue('Y_SECTION',ARecordAna.Section) ;
          Y_.PutValue('Y_LIBELLE',ARecordAna.Libelle) ;
          Y_.PutValue('Y_GENERAL',ARecord.Compte) ;
          if Ecr_.GetValue('E_DEBIT') <> 0 then Y_.PutValue('Y_DEBITDEV',ARecordAna.Montant)
                                           else Y_.PutValue('Y_CREDITDEV',ARecordAna.Montant) ;
          Y_.PutValue('Y_DEVISE',V_PGI.DevisePivot) ;
          Y_.PutValue('Y_JOURNAL', ARecord.CodeJournal) ;
          Y_.PutValue('Y_DATECOMPTABLE',ARecord.wDate) ;
        end ;
      end ;*)
    end ;

    if Visu then
    begin
      ImVisuEcriture(ToutesLesTobs) ;
    end
    else
    begin
      for i:=0 to ToutesLesTobs.Detail.Count-1 do
      begin
        E_s:=ToutesLesTobs.Detail[i] ;
        //E_s.SaveToFile('C:\IntegrS1.txt',true,true,true) ;
        Okok:=false ;
        io:=TIOS1.create ;
        try
          io.IsImport:=True ;
          io.ArreterSiErreur:=ArreterSiErreur ;
          //io.debugIO:=True ;
          Okok:=(io.DispatchInputTob(E_S)=0) ;
        finally
          if not Okok then PgiBox(io.ErreurIO.Libelle,'Erreur') ;
          io.Free ;
        end;
        if Okok and (fTypeEcr=toEcheance) then MajTableEcheance;
        if ArreterSiErreur then break ;
      end ;
    end ;
  finally
    ToutesLesTobs.Free ;
  end ;
end;

procedure TIntegration.MajTableEcheance;
var i : integer;
    Q , QImmo: TQuery;
begin
  if fListeImmo<>nil then
  begin
    InitMove (fListeImmo.Count,'');
    for i := 0 to fListeImmo.Count-1 do
    begin
      MoveCur(False);
      MarqueIntegreEch(fListeImmo[i]) ;
    end;
    FiniMove;
  end
  else
  begin
    QImmo := OpenSQL('SELECT I_IMMO,I_NATUREIMMO,I_ETAT FROM IMMO WHERE (I_NATUREIMMO="LOC" OR I_NATUREIMMO="CB") AND (I_ETAT <> "FER")',True);
    if not QImmo.Eof then
    begin
      InitMove (QImmo.RecordCount,'');
      QImmo.First;
      While not QImmo.Eof do
      begin
        MoveCur(False);
        MarqueIntegreEch(QImmo.FindField('I_IMMO').AsString) ;
        QImmo.Next;
        Ferme (Q);
      end;
    end;
    Ferme (QImmo);
    FiniMove;
  end;
end;

procedure TIntegration.MarqueIntegreEch(CodeImmo: string) ;
var Q: TQuery ;
begin
  Q := OpenSQL ('SELECT * FROM IMMOECHE WHERE IH_IMMO="'+CodeImmo+'"'+
  ' AND IH_DATE>="'+USDate(DateDebutEch)+'" AND IH_DATE<="'+USDate(DateFinEch)+'"'+
  ' AND (IH_INTEGREECH="-" OR IH_INTEGREECH="-")' ,False);
  while not Q.Eof do
  begin
    Q.Edit;
    if cbEcrEcheance.Checked then Q.FindField('IH_INTEGREECH').AsString := 'X';
    if cbEcrPaiement.Checked then Q.FindField('IH_INTEGREPAI').AsString := 'X';
    Q.Post;
    Q.Next;
  end;
  Ferme (Q);
end ;

procedure TIntegration.JBanqueChange(Sender: TObject);
begin
  if JBanque.Value<>'' then CBanque.Text := GetCompteBanqueContrepartie;
end;

function TIntegration.ControleZoneEcheance : boolean;
begin
  result := False ;
  if not (IsValidDate (DateDebutEch.Text)) then HM.Execute (21,Caption,'')
  else if not (IsValidDate (DateFinEch.Text)) then HM.Execute (22,Caption,'')
  else if not cbEcrEcheance.Checked and not cbEcrPaiement.Checked then HM.Execute (24,Caption,'')
  else if cbEcrEcheance.Checked and (JEcheance.Value='') then HM.Execute (16,Caption,'')
  else if cbEcrEcheance.Checked and not presence('GENERAUX','G_GENERAL',CTVA.Text) then HM.Execute (17,Caption,'')
  else if cbEcrPaiement.Checked and (JBanque.Value='') then HM.Execute (15,Caption,'')
  else if cbEcrPaiement.Checked and not presence('GENERAUX','G_GENERAL',CBanque.Text)
       and (GetCompteBanqueContrepartie<>CBanque.Text) then HM.Execute (18,Caption,'')
  else result := True;
end;

function TIntegration.ControleZoneDotation : boolean;
begin
  result:=false ;
  if (CodeJournal.Value='') then HM.Execute (10,Caption,'')
  else if not (IsValidDate (DateCalcul.Text)) then HM.Execute (19,Caption,'')
  else if not (IsValidDate (DateGeneration.Text)) then HM.Execute (20,Caption,'')
  else if ((StrToDate(DateCalcul.Text)>VHImmo^.Encours.Fin) or (StrToDate(DateCalcul.Text)<VHImmo^.Encours.Deb)
     or (StrToDate(DateGeneration.Text)> VHImmo^.Encours.Fin) or (StrToDate(DateGeneration.Text)<VHImmo^.Encours.Deb)) then
     HM.Execute(12,Caption,'')
  else
    result := True;
end;

function TIntegration.GetCompteBanqueContrepartie : string;
var Q : TQuery;
begin
  result := '';
  Q := OpenSQL ('SELECT J_CONTREPARTIE FROM JOURNAL WHERE J_JOURNAL="'+JBanque.Value+'"',True);
  if not Q.Eof then result := Q.FindField('J_CONTREPARTIE').AsString;
  Ferme (Q);
end;

procedure TIntegration.bZoomClick(Sender: TObject);
begin
  if (fTypeEcr=toDotation) then
  begin
    if not ControleZoneDotation then exit;
    PrepareEcrituresDot ;
  end
  else
  begin
    if not ControleZoneEcheance then exit;
    PrepareEcrituresEch ;
  end ;

  IntegrationCompta(true) ;
end;

procedure TIntegration.ToolWindow971Resize(Sender: TObject);
begin
  Bzoom.Left:=2 ;
end;

end.

