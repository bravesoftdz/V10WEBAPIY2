unit ImMutati;

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
  ImAncOpe,
  HSysMenu,
  hmsgbox,
  StdCtrls,
  ComCtrls,
  HRichEdt,
  HRichOLE,
  Mask,
  Hctrls,
  HTB97,
  ExtCtrls,
  HPanel,
  OpEnCour,
  utob,
  {$IFDEF EAGLCLIENT}
  {$ELSE}
  db,
  {$IFNDEF DBXPRESS}dbtables,{$ELSE}uDbxDataSet,{$ENDIF}
  {$ENDIF}
  Hent1,
  ImEnt,
{$IFNDEF CMPGIS35}
  AMSYNTHESEDPI_TOF,
{$ENDIF}
  LookUp;

type
  TFMutation = class(TFAncOpe)
    HM2: THMsgBox;
    bCompteComptable: TCheckBox;
    COMPTEIMMO: THCritMaskEdit;
    CODEIMMO: TEdit;
    bCodeImmo: TCheckBox;
    procedure bCompteComptableClick(Sender: TObject);
    procedure bCodeImmoClick(Sender: TObject);
    procedure COMPTEIMMOElipsisClick(Sender: TObject);
    procedure COMPTEIMMOEnter(Sender: TObject);
    procedure COMPTEIMMOExit(Sender: TObject);
    procedure FormCreate(Sender: TObject);
  private
    { Déclarations privées }
    fDateOp : TDateTime;
    fNouveauCompte, fNouveauCodeImmo : string;
    fCurCompte : string;
    fCpteSup,fCpteInf : string;
    procedure EnregMutation ;
  public
    { Déclarations publiques }
    procedure InitZones;override;
    function ControleZones : boolean;override;
  end;

function ExecuteMutation(Code : string) : TModalResult;
function AnnuleMutation(LogMutation : TLogMutation; TabMessErreur : THMsgBox) : string ;
procedure MajComptesAssocies(CodeImmo,CompteImmo:string);

implementation

uses Outils,ImOutGen,
{$IFDEF SERIE1}
  Ut2Points, uterreur,
{$ELSE}
{$ENDIF}
IMMOCPTE_TOM ;

{$R *.DFM}

function ExecuteMutation(Code : string) : TModalResult;
var  FMutation: TFMutation;
begin
  FMutation:=TFMutation.Create(Application) ;
  FMutation.fCode:=Code ;
  FMutation.fProcEnreg := FMutation.EnregMutation;
  try
    FMutation.ShowModal ;
  finally
    Result := FMutation.ModalResult;
    FMutation.Free ;
  end ;
end;

procedure TFMutation.InitZones ;
begin
  inherited;
  if (fNature <> 'CB') and (fNature <> 'LOC') then
  begin
    Caption := HM2.Mess[7];
    bCompteComptable.Caption:=HM2.Mess[5];
  end
  else if (fNature = 'CB') then
  begin
    Caption := HM2.Mess[8];
    bCompteComptable.Caption:=HM2.Mess[6];
    laI_COMPTEIMMO.Caption := fCompteCharge;
  end
  else if (fNature = 'LOC') then
  begin
    Caption := HM2.Mess[10];
    bCompteComptable.Caption:=HM2.Mess[6];
    laI_COMPTEIMMO.Caption := fCompteCharge;
  end;
  CODEIMMO.Text := '';
  COMPTEIMMO.Text := '';
  if fNature = 'CB' then
  begin fCpteSup := VHImmo^.CpteCBSup; fCpteInf := VHImmo^.CpteCBInf;end
  else if fNature = 'LOC' then
  begin fCpteSup := VHImmo^.CpteLocSup; fCpteInf := VHImmo^.CpteLocInf;end
  else  if fNature = 'PRO' then
  begin fCpteSup := VHImmo^.CpteImmoSup; fCpteInf := VHImmo^.CpteImmoInf;end
  else if fNature = 'FI' then
  begin fCpteSup := VHImmo^.CpteFinSup; fCpteInf := VHImmo^.CpteFinInf;end;
end ;

procedure TFMutation.EnregMutation ;
var TLog : TOB;
    sSet,sWhere : string;
begin
  // Création nouvel enregistrement ImmoLog
  TLog := TOB.Create ('IMMOLOG',nil,-1);
  try
    TLog.PutValue('IL_BLOCNOTE',IL_BLOCNOTE.LinesRTF.Text);
    TLog.PutValue('IL_TYPEOP','MUT');
    if bCodeImmo.Checked then TLog.PutValue('IL_CODEMUTATION',FCode)
    else TLog.PutValue('IL_CODEMUTATION','');
    if bCompteComptable.Checked then
    begin
      if (fNature = 'PRO') or (fNature = 'FI') then
        TLog.PutValue('IL_CPTEMUTATION',fCompteImmo)
      else TLog.PutValue('IL_CPTEMUTATION',fCompteCharge);
    end
    else TLog.PutValue('IL_CPTEMUTATION','');
    TLog.PutValue('IL_IMMO',FCode);
    TLog.PutValue('IL_DATEOP',FDateOp );
    TLog.PutValue('IL_ORDRE',fOrdre);
    TLog.PutValue('IL_ORDRESERIE',fOrdreS);
    TLog.PutValue('IL_PLANACTIFAV',fPlanActif);
    TLog.PutValue('IL_PLANACTIFAP',fPlanActif);
    TLog.PutValue('IL_LIBELLE',RechDom('TIOPEAMOR', TLog.GetValue('IL_TYPEOP'), FALSE)+' '+DateToStr(FDateOp));
    TLog.PutValue('IL_TYPEMODIF',AffecteCommentaireOperation('MUT'));
    TLog.InsertDB(nil);
  finally
    TLog.Free;
  end;
  if bCodeImmo.Checked then
  begin
    ExecuteSQL('UPDATE IMMOLOG SET IL_IMMO="'+FNouveauCodeImmo+'" WHERE IL_IMMO="'+FCode+'"') ;
    ExecuteSQL('UPDATE IMMOECHE SET IH_IMMO="'+FNouveauCodeImmo+'" WHERE IH_IMMO="'+FCode+'"') ;
    ExecuteSQL('UPDATE IMMOAMOR SET IA_IMMO="'+FNouveauCodeImmo+'" WHERE IA_IMMO="'+FCode+'"');
    ExecuteSQL('UPDATE IMMOUO SET IUO_IMMO="'+FNouveauCodeImmo+'" WHERE IUO_IMMO="'+FCode+'"');
    sSet := 'I_IMMO="'+FNouveauCodeImmo+'", I_CHANGECODE="'+FCode+'"';
    sWhere := 'I_IMMO="'+FCode+'"';
  end else FNouveauCodeImmo:=FCode;
  if bCompteComptable.Checked then
  begin
    if (fNature = 'PRO') or (fNature = 'FI') then
    begin
      ExecuteSQL('UPDATE IMMOAMOR SET IA_COMPTEIMMO="'+FNouveauCompte+'" WHERE IA_IMMO="'+FNouveauCodeImmo+'"');
      if sSet='' then sSet := 'I_COMPTEIMMO="'+FNouveauCompte+'", I_COMPTEREF="'+FNouveauCompte+'"'
      else sSet := sSet + ',I_COMPTEIMMO="'+FNouveauCompte+'",I_COMPTEREF="'+FNouveauCompte+'"';
      sWhere := 'I_IMMO="'+fCode+'"';
    end
    else
    begin
      if sSet='' then sSet := 'I_COMPTELIE="'+FNouveauCompte+'",I_COMPTEREF="'+FNouveauCompte+'"'
      else sSet := sSet +',I_COMPTELIE="'+FNouveauCompte+'",I_COMPTEREF="'+FNouveauCompte+'"';
      sWhere := 'I_IMMO="'+fCode+'"';
    end;
    if ExecuteSQL ('UPDATE IMMO SET '+sSet+',I_DATEMODIF="'+UsTime(NowH)+'" WHERE '+sWhere+' AND I_DATEMODIF="'+UsTime(fSaveDate)+'"') < 1
    then V_PGI.IOError := oeSaisie;
    if FNature = 'PRO' then MajComptesAssocies (FNouveauCodeImmo,FNouveauCompte);

    // ajout mbo - FQ 13234 - taxe pro à 0 si ancien cpte = immo corporelle
    // et nouveau compte = incorporelle

    if (copy(FNouveauCompte, 1, 2) = '20') AND (copy(FCompteImmo, 1, 2) <> '20') THEN
    begin
       if ExecuteSQL ('UPDATE IMMO SET '+sSet+',I_BASETAXEPRO=0 WHERE '+sWhere+' AND I_DATEMODIF="'+UsTime(NowH)+'"') < 1
        then V_PGI.IOError := oeSaisie;
    end;
  end
  else if bCodeImmo.Checked then
  begin
    if ExecuteSQL ('UPDATE IMMO SET '+sSet+',I_DATEMODIF="'+UsTime(NowH)+'" WHERE '+sWhere+' AND I_DATEMODIF="'+UsTime(fSaveDate)+'"') < 1
    then V_PGI.IOError := oeSaisie;
  end;

  //Tga 28/06/2006 Maj Immomvtd
{$IFNDEF CMPGIS35}
  IF FCode<>FNouveauCodeImmo Then
    AM_MAJ_IMMOMVTD('M',FCode,FNouveauCodeImmo,0);
{$ENDIF}

  CocheChampOperation (nil,fNouveauCodeImmo,'I_OPEMUTATION');
end;

procedure TFMutation.bCompteComptableClick(Sender: TObject);
begin
  inherited;
  COMPTEIMMO.enabled := bCompteComptable.Checked;
  if bCompteComptable.Checked then FocusControl (COMPTEIMMO);
end;

procedure TFMutation.bCodeImmoClick(Sender: TObject);
begin
  inherited;
  CODEIMMO.enabled := bCodeImmo.Checked;
  if bCodeImmo.Checked then
  begin
    CODEIMMO.Text:=NouveauCodeImmo;
    FocusControl (CODEIMMO);
  end
  else CODEIMMO.Text:='';
end;

function TFMutation.ControleZones : boolean;
var CpteAss : TCompteAss;
{$IFDEF SERIE1}
  Err : TPGIErr ; //XMG 24/10/02
{$ENDIF}
begin
  result := inherited ControleZones;
  if result = false then exit;
  fDateOp:=StrToDate(DATEOPE.Text) ;
  fNouveauCompte:=COMPTEIMMO.Text;
  fNouveauCodeImmo:=CODEIMMO.Text;
  ModalResult := mrYes;
  if (COMPTEIMMO.Text='') and (CODEIMMO.Text='') then
  begin
    PGIBox('Mutation impossible.#10#13Vous devez renseigner un nouveau compte ou un nouveau code.', Caption);
    ModalResult := mrNone; Result := false;
    FocusControl (DATEOPE); exit;
  end;
  if (bCompteComptable.Checked) and (COMPTEIMMO.Text='') then
  begin
     HM2.execute(0,Caption,'');
     ModalResult := mrNone; Result := false;
     FocusControl (COMPTEIMMO); exit;
  end;
  if (bCodeImmo.Checked) and (CODEIMMO.Text='') then
  begin
     HM2.execute(1,Caption,'');
     ModalResult := mrNone; Result := false;
     FocusControl (CODEIMMO);exit;
  end;
  if bCodeImmo.Checked then
  begin
//XMG 24/10/02 début
{$IFDEF SERIE1}
  if not Testecode(fNouveauCodeImmo,Err) then
    Begin
    PGIBox(Err.Libelle,Caption) ;
    FocusControle(CODEIMMO) ;
    ModalResult:=mrNone ;
    Result:=FALSE ; exit ;
    End ;
{$ENDIF}
//XMG 24/10/02 fin
    if ExisteCodeImmo(fNouveauCodeImmo) then
    begin
      HM2.execute(2,Caption,'') ;fNouveauCodeImmo:='' ;
      FocusControl(CODEIMMO);CODEIMMO.Text:='' ;
      // FQ 14523 ajout mbo 29.08.05
      ModalResult:=mrNone ;
      Result := false ;  exit;
    end;
  end;
  if (bCodeImmo.Checked) or (bCompteComptable.Checked) then
  begin
    if bCompteComptable.Checked then
    begin
      if not Presence('GENERAUX','G_GENERAL',COMPTEIMMO.Text) then
      begin
        HM2.Execute(3,Caption,'');FocusControl (COMPTEIMMO);
        ModalResult := mrNone;Result := false; exit;
      end;
      if (COMPTEIMMO.Text<fCpteInf) or (COMPTEIMMO.Text>fCpteSup)then
      begin
        HM2.Execute(9,Caption,'');FocusControl (COMPTEIMMO);
        ModalResult := mrNone;Result := false; exit;
      end;
      if (not bCodeImmo.Checked) and ((fNature='PRO') or (fNature='FI')) and (COMPTEIMMO.Text=fCompteImmo) then
      begin
        HM2.Execute(4,Caption,'');
        ModalResult := mrNone;Result := false; exit;
      end;
      if (not bCodeImmo.Checked) and ((fNature='CB') or (fNature='LOC')) and (COMPTEIMMO.Text=fCompteCharge) then
      begin
        HM2.Execute(4,Caption,'');
        ModalResult := mrNone;Result := false; exit;
      end;
      if fNature = 'PRO' then // Comptes associés uniquement en pleine propriété
      begin
        RecupereComptesAssocies(nil,fNouveauCompte,CpteAss);
        Result := VerifEtCreationComptesAssocies (CpteAss);
        if Result = False then
        begin
          ModalResult := mrNone; exit;
        end;
      end;
    end;
  end;
  result := true;
end;

function AnnuleMutation(LogMutation : TLogMutation; TabMessErreur : THMsgBox) : string;
var NatureImmo,AncienCode,CodeCourant,AncienCompte : string; QueryImmo : TQuery;
begin
  result := LogMutation.Lequel;
  QueryImmo:=OpenSQL('SELECT * FROM IMMO WHERE I_IMMO="'+LogMutation.Lequel+'"', FALSE) ;
  QueryImmo.Edit ;
  NatureImmo:='';
  AncienCode:=LogMutation.CodeMutation;
  if IsOpeEnCours(nil,AncienCode,false) then
  begin
    TabMessErreur.execute(22,'','');
    V_PGI.IoError:=oeUnknown ;
    exit;
  end;
  CodeCourant:=LogMutation.Lequel;
  AncienCompte:=LogMutation.CpteMutation;
  if AncienCode<>'' then
  begin
    ExecuteSQL('UPDATE IMMOAMOR SET IA_IMMO="'+AncienCode+'" WHERE IA_IMMO="'+CodeCourant+'"');
    ExecuteSQL('UPDATE IMMOECHE SET IH_IMMO="'+AncienCode+'" WHERE IH_IMMO="'+CodeCourant+'"') ;
    ExecuteSQL('UPDATE IMMOLOG SET IL_IMMO="'+AncienCode+'" WHERE IL_IMMO="'+CodeCourant+'"') ;
    ExecuteSQL('UPDATE IMMOUO SET IUO_IMMO="'+AncienCode+'" WHERE IUO_IMMO="'+CodeCourant+'"') ;
    QueryImmo.FindField('I_IMMO').AsString:=AncienCode ;
    QueryImmo.FindField('I_CHANGECODE').AsString := '';
    result := AncienCode;
  end else AncienCode := CodeCourant;
  if AncienCompte <>'' then
  begin
    ExecuteSQL('UPDATE IMMOAMOR SET IA_COMPTEIMMO="'+AncienCompte+'" WHERE IA_IMMO="'+AncienCode+'"') ;
    NatureImmo := QueryImmo.FindField('I_NATUREIMMO').AsString;
    if (NatureImmo = 'PRO') or (NatureImmo = 'FI') then
      QueryImmo.FindField('I_COMPTEIMMO').AsString:=AncienCompte
    else QueryImmo.FindField('I_COMPTELIE').AsString:=AncienCompte;
    QueryImmo.FindField('I_COMPTEREF').AsString:=AncienCompte;
    QueryImmo.FindField('I_CHANGECODE').AsString := '';
  end;

  //Tga 28/06/2006 Maj Immomvtd
{$IFNDEF CMPGIS35}
  IF LogMutation.Lequel<>AncienCode Then
    AM_MAJ_IMMOMVTD('M',LogMutation.Lequel,AncienCode,0);
{$ENDIF}

  MajOpeEnCoursImmo (QueryImmo, 'I_OPEMUTATION','MUT','-');
  QueryImmo.Post ;
  Ferme(QueryImmo);
  if NatureImmo = 'PRO'  then MajComptesAssocies(AncienCode,AncienCompte);
end;

procedure MajComptesAssocies(CodeImmo,CompteImmo:string);
var CAss: TCompteAss;
begin
  RecupereComptesAssocies(nil,CompteImmo,CAss) ;
  if VerifEtCreationComptesAssocies (CAss) then
    InsertOrUpdateLesComptesAssocies('IMMO','I_',CodeImmo,CAss,false) ;
end;

procedure TFMutation.COMPTEIMMOElipsisClick(Sender: TObject);
var sWhere : string;
begin
  inherited;
  sWhere := 'G_GENERAL<="'+fCpteSup+'" AND G_GENERAL>="'+fCpteInf+'"';
  {$IFDEF SERIE1}
  LookUpList (TControl (Sender),TraduireMemoire('Comptes'),'GENERAUX','G_GENERAL','G_LIBELLE',sWhere,'G_GENERAL',True,21) ;
  {$ELSE}
  LookUpList (TControl (Sender),TraduireMemoire('Comptes'),'GENERAUX','G_GENERAL','G_LIBELLE',sWhere,'G_GENERAL',True,1) ;
  {$ENDIF}
end;

procedure TFMutation.COMPTEIMMOEnter(Sender: TObject);
begin
  inherited;
  fCurCompte := COMPTEIMMO.Text;
end;

procedure TFMutation.COMPTEIMMOExit(Sender: TObject);
var Compte : string;
begin
  inherited;
  Compte := COMPTEIMMO.Text;
  if Compte = fCurCompte then exit;
  if Compte = '' then exit;
  Compte := ImBourreEtLess ( Compte,ImGeneTofb);
  if Presence('GENERAUX','G_GENERAL',Compte) then
    COMPTEIMMO.Text := Compte;
end;

procedure TFMutation.FormCreate(Sender: TObject);
begin
  inherited;
{$IFDEF SERIE1}
HelpContext:=511030 ;
{$ELSE}
HelpContext:=2111100 ;
{$ENDIF}

end;

end.
