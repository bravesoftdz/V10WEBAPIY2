unit ImLevOpt;

// 05/06/1999 - CA - Prise en compte I_COMPTEREF
// CA - 08/07/1999 - Mise à jour I_ETAT suite éclatement
// CA - 18/08/2000 - Recalcul du taux si durée change.
// MBO - 19/01/2006 - FQ 17339 - stockage de la nouvelle date fin de contrat dans
//                    i_dateFINcb
// BTY - 12/06 FQ 19280 - Stocker l'ancienne date fin de contrat pour pouvoir la restituer en annulation d'op
// BTY - 05/07 FQ 19922 - Mode NAM bloque en entrée opération + autoriser immo générée NAM

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
  db,
  {$IFNDEF EAGLCLIENT}
  {$IFNDEF DBXPRESS}dbtables,{$ELSE}uDbxDataSet,{$ENDIF}
  {$ENDIF}
  OpEnCour,
  Hent1,
  UTOB,
  DBCtrls,
  HDB,
  ImEnt,
  LookUp,
  ImouTGen,
  IMMOCPTE_TOM, Spin;

type
  TFLevOpt = class(TFAncOpe)
    HM2: THMsgBox;
    HLabel4: THLabel;
    CODEIMMO: TEdit;
    HLabel5: THLabel;
    DESIGNATION: TEdit;
    COMPTEIMMO: THCritMaskEdit;
    HLabel2: THLabel;
    HLabel9: THLabel;
    MONTANTLEVEE: THNumEdit;
    GroupBox4: TGroupBox;
    HLabel10: THLabel;
    HLabel11: THLabel;
    HLabel12: THLabel;
    HLabel13: THLabel;
    TAUXECO: THNumEdit;
    METHODEAMOR: THValComboBox;
    DureeUtil: TSpinEdit; // FQ 19922
    procedure OnChangeMethode(Sender: TObject);
    procedure COMPTEIMMOExit(Sender: TObject);
    procedure COMPTEIMMOElipsisClick(Sender: TObject);
    procedure FormCreate(Sender: TObject);
  private
    { Déclarations privées }
    fCodeDepotGarantie : string;
    procedure EnregistreLeveeOption;
    procedure CessionDepotGarantie;
    function MajComptesAssocies (leCompte : string;var CAss : TCompteAss) : boolean;
  public
    { Déclarations publiques }
    procedure InitZones;override;
    function ControleZones : boolean;override;
  end;

function ExecuteLeveeOption (Code : string) : TModalResult;
procedure AnnuleLeveeOption (LogLeveeOpt : TLogLeveeOpt ;TabMessErreur : THMsgBox);

implementation

uses  Outils,
      ImPlan,
      ImSortie,
      ImOuPlan,
      IMMO_TOM
      {$IFDEF SERIE1}
      , Ut2Points
      , uterreur
      {$ENDIF}
      ;

{$R *.DFM}

function ExecuteLeveeOption (Code : string) : TModalResult;
var FLevOpt: TFLevOpt;
begin
  FLevOpt := TFLevOpt.Create(Application);
  FLevOpt.fCode := Code;
  FLevOpt.fProcEnreg := FLevOpt.EnregistreLeveeOption;
  try
    FLevOpt.ShowModal;
  finally
    result := FLevOpt.ModalResult;
    FLevOpt.Free;
  end;
end;

{
procedure TFLevOpt.EnregistreLeveeOption;
var QueryW,QueryR : TQuery;
    bErreurAss : boolean;
    CAss : TCompteAss;
begin
  bErreurAss := false;
  QueryR:=OpenSQL('SELECT * FROM IMMO WHERE I_IMMO="'+fCode+'"', FALSE);
  if not QueryR.EOF then
  begin
   fCodeDepotGarantie := QueryR.FindField('I_IMMOLIEGAR').AsString;
   QueryW:=OpenSQL('SELECT * FROM IMMO WHERE I_IMMO="'+W_W+'"', FALSE);
   QueryW.Insert ; InitNew(QueryW) ;
   QueryW.FindField('I_IMMO').AsString:=CODEIMMO.Text;
   QueryW.FindField('I_ETAT').AsString:='OUV';
   QueryW.FindField('I_CHANGECODE').AsString:=fCode; //18/01/99
   QueryW.FindField('I_NATUREIMMO').AsString:='PRO';
   QueryW.FindField('I_LIBELLE').AsString:=DESIGNATION.Text;
   QueryW.FindField('I_MONTANTHT').AsFloat := MONTANTLEVEE.Value;
   QueryW.FindField('I_VALEURACHAT').AsFloat := MONTANTLEVEE.Value;
   QueryW.FindField('I_QUANTITE').AsFloat:=QueryR.FindField('I_QUANTITE').AsFloat;
   QueryW.FindField('I_TVARECUPERABLE').AsFloat:=0;
   QueryW.FindField('I_TVARECUPEREE').AsFloat:=0;
//12/03/99
   QueryW.FindField('I_REPRISEECO').AsFloat := 0;;
   QueryW.FindField('I_REPRISEFISCAL').AsFloat := 0;;
//   QueryW.FindField('I_BASEECO').AsFloat := QueryR.FindField('I_RESIDUEL').AsFloat;
//   QueryW.FindField('I_BASEFISC').AsFloat := QueryR.FindField('I_RESIDUEL').AsFloat;
   QueryW.FindField('I_BASEECO').AsFloat := MONTANTLEVEE.Value;
   QueryW.FindField('I_BASEFISC').AsFloat := MONTANTLEVEE.Value;
//12/03/99
   QueryW.FindField('I_BASETAXEPRO').AsFloat:=QueryR.FindField('I_MONTANTHT').AsFloat;
   QueryW.FindField('I_METHODEECO').AsString:=METHODEAMOR.Value;
   QueryW.FindField('I_DUREEECO').AsInteger:=StrToInt(DUREEUTIL.Text);
   QueryW.FindField('I_TAUXECO').AsFloat:=TAUXECO.Value;
   QueryW.FindField('I_METHODEFISC').AsString:=QueryR.FindField('I_METHODEFISC').AsString;
   QueryW.FindField('I_DUREEFISC').AsInteger:=QueryR.FindField('I_DUREEFISC').AsInteger;
   QueryW.FindField('I_TAUXFISC').AsFloat:=QueryR.FindField('I_TAUXFISC').AsFloat;
   QueryW.FindField('I_QUALIFIMMO').AsString := QueryR.FindField('I_QUALIFIMMO').AsString;
   QueryW.FindField('I_TABLE0').AsString := QueryR.FindField('I_TABLE0').AsString;
   QueryW.FindField('I_TABLE1').AsString := QueryR.FindField('I_TABLE1').AsString;
   QueryW.FindField('I_COMPTEIMMO').AsString := COMPTEIMMO.Text;
   QueryW.FindField('I_COMPTEREF').AsString := COMPTEIMMO.Text;
   QueryW.FindField('I_DATEPIECEA').AsDateTime := StrToDate(DATEOPE.Text);
   QueryW.FindField('I_DATEAMORT').AsDateTime := StrToDate(DATEOPE.Text);
   QueryW.FindField('I_CODEPOSTAL').AsString := QueryR.FindField('I_CODEPOSTAL').AsString;
   QueryW.FindField('I_VILLE').AsString := QueryR.FindField('I_VILLE').AsString;
   QueryW.FindField('I_PAYS').AsString := QueryR.FindField('I_PAYS').AsString;
   QueryW.FindField('I_ETABLISSEMENT').AsString := QueryR.FindField('I_ETABLISSEMENT').AsString;
   QueryW.FindField('I_LIEUGEO').AsString := QueryR.FindField('I_LIEUGEO').AsString;
   if MajComptesAssocies (QueryW.FindField('I_COMPTEIMMO').AsString,Cass) then
   begin
     QueryW.FindField('I_COMPTEAMORT').AsString := CAss.Amort;
     QueryW.FindField('I_COMPTEDOTATION').AsString := CAss.Dotation;
     QueryW.FindField('I_COMPTEDEROG').AsString := CAss.Derog;
     QueryW.FindField('I_REPRISEDEROG').AsString := CAss.RepriseDerog;
     QueryW.FindField('I_PROVISDEROG').AsString := CAss.ProvisDerog;
     QueryW.FindField('I_DOTATIONEXC').AsString := CAss.DotationExcep;
     QueryW.FindField('I_VACEDEE').AsString := CAss.VaCedee;
     QueryW.FindField('I_AMORTCEDE').AsString := CAss.AmortCede;
     QueryW.FindField('I_REPEXPLOIT').AsString := CAss.RepriseExploit;
     QueryW.FindField('I_REPEXCEP').AsString := CAss.RepriseExcep;
     QueryW.FindField('I_VAOACEDEE').AsString := CAss.VoaCede;
     QueryW.FindField('I_PLANACTIF').AsInteger:=RecalculSauvePlan(QueryW); //23/02/99
     QueryW.Post;
     Ferme(QueryW);
     ExecuteSQL('UPDATE IMMO SET I_IMMOLIE="'+CODEIMMO.Text+'" WHERE I_IMMO="'+fCode+'"');
     // Création nouvel enregistrement ImmoLog
     QueryW:=OpenSQL('SELECT * FROM IMMOLOG WHERE IL_IMMO="'+W_W+'"', FALSE) ;
     QueryW.Insert ;
     QueryW.FindField('IL_IMMO').AsString:=fCode;
     QueryW.FindField('IL_LIBELLE').AsString:=RechDom('TIOPEAMOR', 'LEV', FALSE)+' '+DateToStr(date);
     QueryW.FindField('IL_TYPEMODIF').AsString:=AffecteCommentaireOperation('LEV'); //15/01/99 EPZ
     QueryW.FindField('IL_DATEOP').AsDateTime:=StrToDate(DATEOPE.Text);
     QueryW.FindField('IL_TYPEOP').AsString:='LEV' ;
     QueryW.FindField('IL_CODECB').AsString:=CODEIMMO.Text;
     TBlobField(QueryW.FindField('IL_BLOCNOTE')).Assign(IL_BLOCNOTE.LinesRTF);
     QueryW.FindField('IL_ORDRE').AsInteger:=fOrdre;
     QueryW.FindField('IL_ORDRESERIE').AsInteger:=fOrdreS;
     QueryW.Post;
     Ferme(QueryW) ;
     Ferme(QueryR) ;
     // 15/01/99
//       ExecuteSQL('UPDATE IMMO SET I_OPELEVEEOPTION="X" WHERE I_IMMO="'+fCode+'"') ;
     QueryR := nil;
     CocheChampOperation (QueryR,fCode,'I_OPELEVEEOPTION');
     // 15/01/99
   end else
   begin
     bErreurAss := True;
     Ferme (QueryW);
   end;
 end
 else Ferme (QueryR);
 if not bErreurAss then CessionDepotGarantie;
end;
}

procedure TFLevOpt.CessionDepotGarantie;
var Q : TQuery;
begin
  if fCodeDepotGarantie <> '' then
  begin
    Q := OpenSQL ('SELECT I_QUANTITE FROM IMMO WHERE I_IMMO="'+fCodeDepotGarantie+'"',True);
    if not Q.Eof then
      if Q.FindField('I_QUANTITE').AsFloat > 0.0 then
      begin
        if HM2.Execute (4,Caption,'') = mrYes then
        begin
          // Cession du dépôt de garantie
          ExecuteCession(fCodeDepotGarantie);
        end
        else exit;
      end;
    Ferme (Q);
  end;
end;

procedure TFLevOpt.InitZones;
begin
  inherited;
  COMPTEIMMO.Text:=fCompteImmo;
  DESIGNATION.Text := fLibelle ;
  MONTANTLEVEE.Value := fResiduel;
  TAUXECO.Value := fTauxEco;
  DUREEUTIL.Value := fDureeEco;
  METHODEAMOR.Value := fMethodeEco;
  CODEIMMO.Text := NouveauCodeImmo;
  DATEOPE.Text := DateToStr(fDateFinContrat);

  // FQ 16291 tga 29/08/2005
  MONTANTLEVEE.masks.PositiveMask := StrfMask(V_PGI.OkDecV,'', True);
  // fin tga 29/08/2005

  onchangeMethode (nil);
end;

function TFLevOpt.ControleZones : boolean;
{$IFDEF SERIE1}
var Err : TPGIErr ; //XMG 24/10/02
    NumMess : integer;
{$ELSE}
var NumMess : integer;
{$ENDIF}
begin
  result := inherited ControleZones;
  if result = false then exit;
  // Contrôle de la date d'opération par rapport à la date de fin de contrat
  if (StrToDate(DATEOPE.Text) <> fDateFinContrat) then
  begin
    if HM2.Execute (6,Caption,'')=mrNo then
    begin result := false;exit;end;
  end;
  // Contrôle du code et du libellé
  if (CODEIMMO.Text = '')  or (DESIGNATION.Text = '') then
  begin
    HM2.Execute (0,Caption,'');
    Result := false; exit;
  end;
//XMG 24/10/02 début
{$IFDEF SERIE1}
  if not Testecode(CODEIMMO.Text,Err) then
    Begin
    PGIBox(Err.Libelle,Caption) ;
    FocusControle(CODEIMMO) ;
    ModalResult:=mrNone ;
    Result:=FALSE ; exit ;
    End ;
{$ENDIF}
//XMG 24/10/02 fin
  // Contrôle du compte d'immobilisation
  if (not ExisteSQL('SELECT G_GENERAL FROM GENERAUX WHERE G_GENERAL="'+COMPTEIMMO.Text+'"')) then
  begin
    HM2.Execute (1,Caption,'');
    Result := false; FocusControl (COMPTEIMMO);exit;
  end;
  // Contrôle du montant
  if (MONTANTLEVEE.Value = 0.0) then
  begin
    HM2.Execute (2,Caption,'');
    Result := false; FocusControl (MONTANTLEVEE);exit;
  end;
  // Controle de la durée d'amortissement
  //if (DUREEUTIL.Value < 12) then
  //begin
  //  HM2.Execute (3,Caption,'');
  //  Result := false; FocusControl (DUREEUTIL);exit;
  //end;
  OnChangeMethode(nil);
  // FQ 19922 Onchangemethode avant de contrôler la durée et non après
  if (not TraiteIntervaleDureeAmort(MethodeAmor.Value,
                                    0,
                                    DureeUtil.Value,
                                    NumMess)) then
  begin
    HM2.Execute (NumMess,Caption,'');
    Result := false; FocusControl (DUREEUTIL);exit;
  end;

  Result := True;
end;

function TFLevOpt.MajComptesAssocies (leCompte : string;var CAss : TCompteAss) : boolean;
BEGIN
//EPZ 09/11/00
  RecupereComptesAssocies(nil, leCompte,CAss);
  Result := VerifEtCreationComptesAssocies (CAss);
end;
{
function TFLevOpt.MajComptesAssocies (var Q : TQuery) : boolean;
Var CAss : TCompteAss;
BEGIN
//EPZ 09/11/00
//  CAss := RecupereComptesAssocies(nil, Q.FindField('I_COMPTEIMMO').AsString);
  RecupereComptesAssocies(nil, Q.FindField('I_COMPTEIMMO').AsString,CAss);
//EPZ 09/11/00
  Q.FindField('I_COMPTEAMORT').AsString := CAss.Amort;
  Q.FindField('I_COMPTEDOTATION').AsString := CAss.Dotation;
  Q.FindField('I_COMPTEDEROG').AsString := CAss.Derog;
  Q.FindField('I_REPRISEDEROG').AsString := CAss.RepriseDerog;
  Q.FindField('I_PROVISDEROG').AsString := CAss.ProvisDerog;
  Q.FindField('I_DOTATIONEXC').AsString := CAss.DotationExcep;
  Q.FindField('I_VACEDEE').AsString := CAss.VaCedee;
  Q.FindField('I_AMORTCEDE').AsString := CAss.AmortCede;
  Q.FindField('I_REPEXPLOIT').AsString := CAss.RepriseExploit;
  Q.FindField('I_REPEXCEP').AsString := CAss.RepriseExcep;
  Q.FindField('I_VAOACEDEE').AsString := CAss.VoaCede;
  Result := VerifEtCreationComptesAssocies (CAss);
end;
}

procedure TFLevOpt.OnChangeMethode(Sender: TObject);
begin
  inherited;
  // FQ 19922
  //if ((IsValidDate(DATEOPE.Text)) and ( Valeur(DUREEUTIL.Text) > 0)) then
  //    TAUXECO.Value := GetTaux(METHODEAMOR.Value,StrToDate(DATEOPE.Text),StrToDate(DATEOPE.Text),StrToInt(DUREEUTIL.Text))

  if MethodeAmor.Value  = 'NAM' then
     begin
     DureeUtil.Enabled := False;
     DureeUtil.Value := 0;
     end
  else
     DureeUtil.Enabled := True;
  if DureeUtil.Text='' then DureeUtil.Value := 1;
  if (IsValidDate(DATEOPE.Text)) then
     TAUXECO.Value := GetTaux(METHODEAMOR.Value,
                              StrToDate(DATEOPE.Text),
                              StrToDate(DATEOPE.Text),
                              DUREEUTIL.Value);
end;


procedure TFLevOpt.COMPTEIMMOExit(Sender: TObject);
begin
  inherited;
end;

// Procedure d'annulation de la levée de l'option
//procedure AnnuleLeveeOption (var CodeCB : string{; QueryLog : TQuery};TabMessErreur : THMsgBox);
procedure AnnuleLeveeOption (LogLeveeOpt : TLogLeveeOpt ;TabMessErreur : THMsgBox);
var Q: TQuery;
    CodeLevee : string;
begin
  Q:=OpenSQL('SELECT * FROM IMMO WHERE I_IMMO="'+LogLeveeOpt.Lequel+'"', FALSE) ;
  Q.Edit ;
  CodeLevee := Q.FindField('I_IMMOLIE').AsString;
  if IsOpeEnCours(nil,CodeLevee,false) then
  begin
    TabMessErreur.execute(23,'','');
    V_PGI.IoError:=oeUnknown ;
    exit;
  end;

  // FQ 17339 - il faut remettre à blanc la date FIN cb dans laquelle on avait
  // stocké la date fin de contrat suite levée d'option
  //Q.FindField ('I_DATEFINCB').AsDateTime := iDate1900;
  // FQ 19280 Restituer l'ancienne date de fin de contrat
  Q.FindField ('I_DATEFINCB').AsDateTime := LogLeveeOpt.DateFinCt;

  ExecuteSQL('DELETE FROM IMMO WHERE I_IMMO="'+CodeLevee+'"');
  ExecuteSQL('DELETE FROM IMMOAMOR WHERE IA_IMMO="'+CodeLevee+'"');
  ExecuteSQL('DELETE FROM IMMOLOG WHERE IL_IMMO="'+CodeLevee+'"');
  Q.FindField('I_IMMOLIE').AsString := '';
  MajOpeEnCoursImmo ( Q,'I_OPELEVEEOPTION', 'LEV', '-');
  Q.Post;
  Ferme(Q);
end;

procedure TFLevOpt.EnregistreLeveeOption;
var bErreurAss : boolean;
    CAss : TCompteAss;
    TQR, TQW, TLog : TOB;
    theQuery : TQuery;
    Infos : TInfoLog;
    NewFinCt : TDateTime;
begin
  bErreurAss := True;
  TQR := TOB.Create('IMMO',nil,-1);
  TQW := TOB.Create('IMMO',nil,-1);
  if TQR.SelectDB('"'+fCode+'"',nil) then
  begin
    fCodeDepotGarantie := TQR.GetValue('I_IMMOLIEGAR');
    TQW.PutValue('I_IMMO',CODEIMMO.Text);
    TQW.PutValue('I_ETAT','OUV');
    TQW.PutValue('I_CHANGECODE',fCode);
    TQW.PutValue('I_NATUREIMMO','PRO');
    TQW.PutValue('I_LIBELLE',DESIGNATION.Text);
    TQW.PutValue('I_MONTANTHT',MONTANTLEVEE.Value);
    TQW.PutValue('I_VALEURACHAT',MONTANTLEVEE.Value);
    TQW.PutValue('I_QUANTITE',TQR.GetValue('I_QUANTITE'));
    TQW.PutValue('I_TVARECUPERABLE',0);
    TQW.PutValue('I_TVARECUPEREE',0);
    TQW.PutValue('I_REPRISEECO',0);
    TQW.PutValue('I_REPRISEFISCAL',0);
    TQW.PutValue('I_BASEECO',MONTANTLEVEE.Value);
    TQW.PutValue('I_BASEFISC',MONTANTLEVEE.Value);
    TQW.PutValue('I_BASETAXEPRO',TQR.GetValue('I_MONTANTHT'));
    TQW.PutValue('I_METHODEECO',METHODEAMOR.Value);
    TQW.PutValue('I_DUREEECO',DUREEUTIL.Value); // FQ 19922 StrToInt(DUREEUTIL.Text));
    TQW.PutValue('I_TAUXECO',TAUXECO.Value);
    TQW.PutValue('I_METHODEFISC',TQR.GetValue('I_METHODEFISC'));
    TQW.PutValue('I_DUREEFISC',TQR.GetValue('I_DUREEFISC'));
    TQW.PutValue('I_TAUXFISC',TQR.GetValue('I_TAUXFISC'));
    TQW.PutValue('I_QUALIFIMMO',TQR.GetValue('I_QUALIFIMMO'));
    TQW.PutValue('I_TABLE0',TQR.GetValue('I_TABLE0'));
    TQW.PutValue('I_TABLE1',TQR.GetValue('I_TABLE1'));
    TQW.PutValue('I_COMPTEIMMO',COMPTEIMMO.Text);
    TQW.PutValue('I_COMPTEREF',COMPTEIMMO.Text);

    TQW.PutValue('I_DATEPIECEA',StrToDate(DATEOPE.Text));
    TQW.PutValue('I_DATEAMORT',StrToDate(DATEOPE.Text));
    TQW.PutValue('I_CODEPOSTAL',TQR.GetValue('I_CODEPOSTAL'));
    TQW.PutValue('I_VILLE',TQR.GetValue('I_VILLE'));
    TQW.PutValue('I_PAYS',TQR.GetValue('I_PAYS'));
    TQW.PutValue('I_ETABLISSEMENT',TQR.GetValue('I_ETABLISSEMENT'));
    TQW.PutValue('I_LIEUGEO',TQR.GetValue('I_LIEUGEO'));
    TQW.PutValue('I_NATUREBIEN',TQR.GetValue('I_NATUREBIEN'));


    //+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
    // tga 25/09/2006
    TQW.PutValue('I_DATEDEBECO',StrToDate(DATEOPE.Text));
    TQW.PutValue('I_DATEDEBFIS',StrToDate(DATEOPE.Text));

//    // TGA 25/04/2006
//    IF METHODEAMOR.Value='DEG' Then
//      Begin
//        TQW.PutValue('I_DATEDEBECO',TQR.GetValue('I_DATEPIECEA'));
//        TQW.PutValue('I_DATEDEBFIS',TQR.GetValue('I_DATEPIECEA'));
//     End
//    Else
//      Begin
//        TQW.PutValue('I_DATEDEBECO',TQR.GetValue('I_DATEAMORT'));
//        TQW.PutValue('I_DATEDEBFIS',TQR.GetValue('I_DATEAMORT'));
//      End;
//
//    IF TQR.GetValue('I_METHODEFISC') ='DEG' Then
//      TQW.PutValue('I_DATEDEBFIS',TQR.GetValue('I_DATEPIECEA'))
//    Else
//      TQW.PutValue('I_DATEDEBFIS',TQR.GetValue('I_DATEAMORT'));

    // fin tga 25/09/2006
    //+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++



    if MajComptesAssocies (TQR.GetValue('I_COMPTEIMMO'),Cass) then
    begin
      TQW.PutValue('I_COMPTEAMORT',CAss.Amort);
      TQW.PutValue('I_COMPTEDOTATION',CAss.Dotation);
      TQW.PutValue('I_COMPTEDEROG',CAss.Derog);
      TQW.PutValue('I_REPRISEDEROG',CAss.RepriseDerog);
      TQW.PutValue('I_PROVISDEROG',CAss.ProvisDerog);
      TQW.PutValue('I_DOTATIONEXC',CAss.DotationExcep);
      TQW.PutValue('I_VACEDEE',CAss.VaCedee);
      TQW.PutValue('I_AMORTCEDE',CAss.AmortCede);
      TQW.PutValue('I_REPEXPLOIT',CAss.RepriseExploit);
      TQW.PutValue('I_REPEXCEP',CAss.RepriseExcep);
      TQW.PutValue('I_VAOACEDEE',CAss.VoaCede);
      TQW.PutValue('I_PLANACTIF',RecalculSauvePlanTOB(TQW));
      TQW.InsertDB(nil);

      ExecuteSQL('UPDATE IMMO SET I_IMMOLIE="'+CheckdblQuote(CODEIMMO.Text)+'" WHERE I_IMMO="'+fCode+'"');

      Infos.TVARecuperable := 0;
      Infos.TVARecuperee := 0 ;
      EnregLogAcquisition(CODEIMMO.Text,StrToDate(DATEOPE.Text), 1,Infos);
      // Création nouvel enregistrement ImmoLog
      TLog := TOB.Create('IMMOLOG',nil,-1);
      TLog.PutValue('IL_IMMO',fCode);
      TLog.PutValue('IL_LIBELLE',RechDom('TIOPEAMOR', 'LEV', FALSE)+' '+DateToStr(date));
      TLog.PutValue('IL_TYPEMODIF',AffecteCommentaireOperation('LEV'));
      TLog.PutValue('IL_DATEOP',StrToDate(DATEOPE.Text));
      TLog.PutValue('IL_TYPEOP','LEV');
      TLog.PutValue('IL_CODECB',CODEIMMO.Text);
      TLog.PutValue('IL_BLOCNOTE',IL_BLOCNOTE.LinesRTF.Text);
      TLog.PutValue('IL_ORDRE',fOrdre);
      TLog.PutValue('IL_ORDRESERIE',fOrdreS);
      // FQ 19280
      TLog.PutValue('IL_DATEOPREELLE', fDateFinContrat);
      TLog.InsertDB(nil);
      TLog.Free;

      theQuery := nil;
      CocheChampOperation (theQuery,fCode,'I_OPELEVEEOPTION');
      // fq 17339
      if (StrToDate(DATEOPE.Text) <> fDateFinContrat) then
      begin
         newFinCt := StrToDate(DATEOPE.Text);
         ExecuteSQL('UPDATE IMMO SET I_DATEFINCB ="' +USDateTime(NewFinCt)+
                    '" WHERE I_IMMO="'+fCode+'"') ;
      end;
      bErreurAss := false;
    end;
  end;
  TQR.Free;
  TQW.Free;
  if not bErreurAss then CessionDepotGarantie;
end;

procedure TFLevOpt.COMPTEIMMOElipsisClick(Sender: TObject);
var stWhere : string;
begin
  inherited;
  stWhere := 'G_GENERAL<="'+VHImmo^.CpteImmoSup+'" AND G_GENERAL>="'+VHImmo^.CpteImmoInf+'" OR '
      +'G_GENERAL<="'+VHImmo^.CpteFinSup+'" AND G_GENERAL>="'+VHImmo^.CpteFinInf+'"';
  LookupList(TControl(Sender),TraduireMemoire('Comptes'),'GENERAUX','G_GENERAL','G_LIBELLE',stWhere,'G_GENERAL', True,1)  ;
end;

procedure TFLevOpt.FormCreate(Sender: TObject);
begin
  inherited;
{$IFDEF SERIE1}
HelpContext:=511030 ;
{$ELSE}
HelpContext:=2111900 ;
{$ENDIF}
end;

end.
