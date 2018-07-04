{***********UNITE*************************************************
Auteur  ...... : Vincent Laroche
Créé le ...... : 16/04/2003
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : CPCFONBMP ()
Mots clefs ... : TOF;CPCFONBMP
*****************************************************************}
Unit CPCFONBMP_TOF ;

Interface

Uses StdCtrls,
     Controls,
     Classes,
{$IFDEF EAGLCLIENT}
     MaineAGL,    // AGLLanceFiche
     eMul,        // TFMul
{$ELSE}
     FE_Main,     // AGLLanceFiche
     db,
     dbtables,
     HDB,         // THDBGrid
     Mul,         // TFMul
{$ENDIF}
     Saisie,      // TrouveEtLanceSaisie
     HQry,        // THQuery
     forms,
     sysutils,
     ComCtrls,
     HCtrls,      // THGrid
     HEnt1,
     HMsgBox,     // Blocage
     SaisUtil,    // TSuiviMP
     MulSMPUtil,  // SmpToStr, StrToSmp
     HTB97,       // TToolBarButton97
     CFONB,       // ExportCFONB
     ed_tools,    // VideListe
     ParamDat,    // ParamDate
     Ent1,        // VH, ExoToDates
     TofVerifRib, // CPLanceFiche_VerifRib
     SaisComm,    // ModifRibOBM
     LettUtil,    // InitTablesLibresTiers
     UTOB,        // TOB
     ParamSoc,		// GetParamSocSecur YMO
     UTOF ;

procedure ExportCFONBMP(smp : TSuiviMP);

Type
  TOF_CPCFONBMP = Class (TOF)
    HM : THMsgBox;
{$IFDEF EAGLCLIENT}
    FListe : THGrid;
{$ELSE}
    FListe : THDBGrid;
{$ENDIF}
    Q : THQuery;
    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnDisplay                ; override ;
    procedure OnClose                  ; override ;
    procedure OnCancel                 ; override ;
    procedure InitMsgBox;

    procedure BOuvrirClick(Sender: TObject);
    procedure FListeDblClick(Sender: TObject);
    procedure BChercheClick(Sender: TObject);

    procedure Auxiliaire1Change(Sender: TObject);
    procedure Auxiliaire2Change(Sender: TObject);
    procedure E_EXERCICEChange(Sender: TObject);
    procedure E_DATECOMPTABLEKeyPress(Sender: TObject; var Key: Char);
    procedure bRibClick(Sender: TObject);
    procedure cRIBClick(Sender: TObject);
    procedure cExportClick(Sender: TObject);
    procedure CTIDTICClick(Sender: TObject);
    procedure BCtrlRibClick(Sender: TObject);
    procedure AuxiElipsisClick(Sender : TObject);
  private
    BanqueGene : String ;
    Function  GetLeOBM : TOBM ;
    Function  ExporteSelection : boolean ;
    Function  CoherBanque ( TL : TList ) : boolean ;
    Procedure InitTitres ;
  public
    SMP : TSuiviMP ;
    IsCli : boolean ;
  end ;

Implementation
uses
  {$IFDEF MODENT1}
  ULibExercice,
  {$ENDIF MODENT1}

  {$IFDEF eAGLCLIENT}
  MenuOLX
  {$ELSE}
  MenuOLG
  {$ENDIF eAGLCLIENT}
  , UTofMulParamGen; {13/04/07 YMO F5 sur Auxiliaire }

procedure ExportCFONBMP(smp : TSuiviMP);
begin
  if Blocage(['nrCloture','nrBatch','nrLettrage'],True,'nrAucun') then Exit;
  AGLLanceFiche('CP','CPCFONBMP','','', SmpToStr(smp));
end;

function SMPClient(smp : TSuiviMP ) : Boolean;
begin
  Result := (smp in [smpEncTraEnc,smpEncTraEsc,smpEncPreBqe]);
end;

procedure TOF_CPCFONBMP.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_CPCFONBMP.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_CPCFONBMP.OnUpdate ;
begin
  Inherited ;
end ;

procedure TOF_CPCFONBMP.OnLoad ;
var
  Q : TQuery;
begin
  Inherited ;
{$IFDEF EAGLCLIENT}
  Q := TFMul(Ecran).Q.TQ;
{$ELSE}
  Q := TFMul(Ecran).Q;
{$ENDIF}
  if (Q <> nil) then SetControlEnabled('BRIB', not (Q.EOF));
end ;

procedure TOF_CPCFONBMP.OnArgument (S : String ) ;
var
  NomFiltre : String;
begin
  Inherited ;
  Q := TFMul(Ecran).Q;
{$IFDEF EAGLCLIENT}
  FListe := THGrid(TFMul(Ecran).FListe);
{$ELSE}
  FListe := THDBGrid(TFMul(Ecran).FListe);
{$ENDIF}

  // Création des contrôles
  HM := THMsgBox.Create(FMenuG);
  InitMsgBox;

  // Récupère les arguments
  SetControlText('XX_WHERE1', 'E_JOURNAL="zzz"');
  SMP := StrToSmp(ReadTokenSt(S));
  IsCli := SMPClient(SMP);

  if IsCli then begin
    Q.Liste := 'CPCFONBMPCLI';
    NomFiltre := 'MULCFONBMPCLI';
    end
  else begin
    if (smp in [smpDecVirInBqe]) then begin
      Q.Liste:='CPCFONBINMPFOU' ;
      NomFiltre:='MULCFONBINMPFOU' ;
      SetControlVisible('BCTRLRIB',False);
      end
    else begin
      Q.Liste:='CPCFONBMPFOU' ;
      NomFiltre:='MULCFONBMPFOU' ;
    end;
  end;
  TFMul(Ecran).FNomFiltre := NomFiltre;

  // Evénements des contrôles
  TToolBarButton97(GetControl('BOUVRIR',True)).OnClick := BOuvrirClick;
{$IFDEF EAGLCLIENT}
  THGrid(GetControl('FLISTE',True)).OnDblClick := FListeDblClick;
{$ELSE}
  THDBGrid(GetControl('FLISTE',True)).OnDblClick := FListeDblClick;
{$ENDIF}
  TToolBarButton97(GetControl('BCHERCHE',True)).OnClick := BChercheClick;
  THEdit(GetControl('AUXILIAIRE1',True)).OnChange := Auxiliaire1Change;
  THEdit(GetControl('E_GENERAL',True)).OnChange := Auxiliaire1Change;
  THEdit(GetControl('AUXILIAIRE2',True)).OnChange := Auxiliaire2Change;
  THValComboBox(GetControl('E_EXERCICE',True)).OnChange := E_EXERCICEChange;
  THEdit(GetControl('E_DATECOMPTABLE',True)).OnKeyPress := E_DATECOMPTABLEKeyPress;
  THEdit(GetControl('E_DATECOMPTABLE_',True)).OnKeyPress := E_DATECOMPTABLEKeyPress;
  THEdit(GetControl('E_DATEECHEANCE',True)).OnKeyPress := E_DATECOMPTABLEKeyPress;
  THEdit(GetControl('E_DATEECHEANCE_',True)).OnKeyPress := E_DATECOMPTABLEKeyPress;
  TToolBarButton97(GetControl('BRIB',True)).OnClick := bRibClick;
  TToolBarButton97(GetControl('BCTRLRIB',True)).OnClick := BCtrlRibClick;
  TCheckBox(GetControl('CRIB',True)).OnClick := cRIBClick;
  TCheckBox(GetControl('CEXPORT',True)).OnClick := cExportClick;
  TCheckBox(GetControl('CTIDTIC',True)).OnClick := CTIDTICClick;

  // FormShow
  SetControlText('E_DEVISE', V_PGI.DevisePivot);
  if (smp in [smpDecVirInBqe]) then begin
    SetControlEnabled('E_DEVISE', True);
    SetControlProperty('BRIB', 'HINT', 'Modifier l''IBAN');
    SetControlText('CRIB', '&IBAN renseigné');
  end;
  if (VH^.CPExoRef.Code<>'') then begin
    SetControlText('E_EXERCICE', VH^.CPExoRef.Code);
    E_EXERCICEChange(Nil);
    SetControlText('E_DATECOMPTABLE', DateToStr(VH^.CPExoRef.Deb));
    SetControlText('E_DATECOMPTABLE_', DateToStr(VH^.CPExoRef.Fin));
    end
  else begin
    SetControlText('E_EXERCICE', VH^.Entree.Code);
    E_EXERCICEChange(Nil) ;
    SetControlText('E_DATECOMPTABLE', DateToStr(V_PGI.DateEntree));
    SetControlText('E_DATECOMPTABLE_', DateToStr(V_PGI.DateEntree));
  end;
  SetControlText('E_DATEECHEANCE', StDate1900);
  SetControlText('E_DATEECHEANCE_',StDate2099);
  InitTablesLibresTiers(TTabSheet(GetControl('PLIBRES',True)));
  SetControlText('XX_WHERE1', '');
  BanqueGene := '';

  if IsCli then begin
    Ecran.Caption := HM.Mess[5]; // Export CFONB des encaissements
    SetControlText('XX_WHEREDC', '');
    Ecran.HelpContext := 7586000;
    SetControlProperty('AUXILIAIRE1', 'DATATYPE', 'tztToutDebit');
    SetControlProperty('AUXILIAIRE2', 'DATATYPE', 'tztToutDebit');
    SetControlText('XX_WHERENATTIERS', 'T_NATUREAUXI="CLI" OR T_NATUREAUXI="AUD"');
    end
  else begin
    Ecran.Caption := HM.Mess[6]; // Export CFONB des décaissements
    SetControlText('XX_WHEREDC', '');
    Ecran.HelpContext := 7595000;
    SetControlProperty('AUXILIAIRE1', 'DATATYPE', 'tztToutCredit');
    SetControlProperty('AUXILIAIRE2', 'DATATYPE', 'tztToutCredit');
    SetControlText('XX_WHERENATTIERS', 'T_NATUREAUXI="FOU" OR T_NATUREAUXI="AUC"');
  end;
  InitTitres ;
  UpdateCaption(Ecran);

  if GetParamSocSecur('SO_CPMULTIERS', false) then
  begin
    THEdit(GetControl('AUXILIAIRE1', true)).OnElipsisClick:=AuxiElipsisClick;
    THEdit(GetControl('AUXILIAIRE2', true)).OnElipsisClick:=AuxiElipsisClick;
  end;
end ;

procedure TOF_CPCFONBMP.OnClose ;
begin
  Inherited ;
  HM.Free;  
end ;

procedure TOF_CPCFONBMP.OnDisplay () ;
begin
  Inherited ;
end ;

procedure TOF_CPCFONBMP.OnCancel () ;
begin
  Inherited ;
end ;

procedure TOF_CPCFONBMP.InitMsgBox;
begin
  HM.Mess.Add('0;?caption?;Export ou émission impossible. Vous n''avez rien sélectionné.;E;O;O;O;');
  HM.Mess.Add('1;?caption?;Confirmez-vous l''export ou l''émission des échéances sélectionnées ?;Q;YN;Y;Y;');
  HM.Mess.Add('2;?caption?;ATTENTION ! Vous avez tout sélectionné. Confirmez-vous l''export ou l''émission ?;W;YN;Y;Y;');
  HM.Mess.Add('3;?caption?;Vous devez sélectionner un modèle de document pour l''émission du bordereau.;W;O;O;O;');
  HM.Mess.Add('Emission de bordereaux');
  HM.Mess.Add('Export CFONB des encaissements');
  HM.Mess.Add('Export CFONB des décaissements');
  HM.Mess.Add('Emission de lettres-BOR');
  HM.Mess.Add('8;?caption?;Vous devez sélectionner un modèle de document.;W;O;O;O;');
  HM.Mess.Add('Modèle');
  HM.Mess.Add('10;?caption?;Vous devez renseigner un compte général.;W;O;O;O;');
  HM.Mess.Add('Emission de lettres-traite');
  HM.Mess.Add('12;?caption?;ATTENTION ! Vous avez des banques différentes. L''export sera incohérent. Confirmez-vous l''export ou l''émission ?;W;YN;Y;Y;');
  HM.Mess.Add('Export CFONB des traites-BOR à l''escompte');
  HM.Mess.Add('Export CFONB des traites-BOR à l''encaissement');
  HM.Mess.Add('Export CFONB des prélèvements');
  HM.Mess.Add('Export CFONB des virements');
  HM.Mess.Add('Export CFONB des BOR');
  HM.Mess.Add('Export CFONB des virements internationaux');
  HM.Mess.Add('19');
end;

function TOF_CPCFONBMP.CoherBanque(TL: TList): boolean;
var
  i : integer ;
  O : TOBM ;
  Jal,OldJal : String3 ;
  Okok : boolean ;
  QQ   : TQuery ;
begin
  if IsCli then begin Result := True; Exit; end;
  Okok := True;
  OldJal := '';
  for i := 0 to TL.Count-1 do begin
    O := TOBM(TL[i]);
    if O=Nil then Break;
    Jal := O.GetMvt('E_JOURNAL');
    if ((OldJal<>'') and (Jal<>OldJal)) then begin Okok := False; Break; end;
    OldJal := Jal;
  end;

  if Not Okok then begin
    // Vous avez des banques différentes. L'export sera incohérent. Confirmez-vous l'export ou l'émission ?
    Okok := (HM.Execute(12,Ecran.Caption,'')=mrYes);
    end
  else if OldJal<>'' then begin
    QQ := OpenSQL('SELECT J_CONTREPARTIE FROM JOURNAL WHERE J_JOURNAL="'+OldJal+'"',True);
    if Not QQ.EOF then BanqueGene := QQ.Fields[0].AsString;
    Ferme(QQ);
  end;
  Result := Okok;
end;

function TOF_CPCFONBMP.ExporteSelection: boolean;
var
  i,NbLig,ii : integer ;
  TL : TList ;
  O       : TOBM ;
  Inutile : TMSEncaDeca ;
{$IFDEF EAGLCLIENT}
  Fiche : TFMul;
{$ENDIF}  
begin
  Result := False;
  Fillchar(Inutile,SizeOf(Inutile),#0) ;

  if Not FListe.AllSelected then begin
    NbLig := Fliste.NbSelected;
    if (NbLig <= 0) then begin
      HM.Execute(0,Ecran.Caption,''); // Export ou émission impossible. Vous n'avez rien sélectionné.
      Exit;
    end;

    // Confirmez-vous l'export ou l'émission des échéances sélectionnées ?
    if HM.Execute(1,Ecran.Caption,'')<>mrYes then Exit;

    {$IFDEF EAGLCLIENT}
    Fiche := TFMul(Ecran);
    if Fiche.bSelectAll.Down then
      if not Fiche.FetchLesTous then
      begin
        Fiche.bSelectAllClick(nil);
        Fiche.bSelectAll.Down := False;
        Exit;
      end;
    {$ENDIF}

    TL := Tlist.Create;
    for i:=0 to NbLig-1 do begin
      FListe.GotoLeBookMark(i);
      O := GetLeOBM;
      if O<>Nil then TL.Add(O);
    end;
    end
  else begin
    // Vous avez tout sélectionné. Confirmez-vous l''export ou l''émission ?
    if HM.Execute(2,Ecran.Caption,'')<>mrYes then Exit ;

    TL := TList.Create;
    Q.First ;
    ii:=0 ;
    While ((Not Q.EOF) and (ii<1000)) do begin
      O := GetLeOBM;
      if O<>Nil then TL.Add(O);
      Q.Next;
      Inc(ii);
    end;
  end;

  if CoherBanque(TL) then ExportCFONB(isCli,BanqueGene,'','DEM',TL,smp);
  VideListe(TL);
  TL.Free;
  Result := True;
  BChercheClick(Nil);
end;

function TOF_CPCFONBMP.GetLeOBM: TOBM;
var
  Q1 : TQuery ;
  O  : TOBM ;
begin
{$IFDEF EAGLCLIENT}
  Q.TQ.Seek(FListe.Row-1);
{$ENDIF}
  O:=Nil;
  Q1:=OpenSQL('SELECT * FROM ECRITURE WHERE E_JOURNAL="' + Q.FindField('E_JOURNAL').AsString + '"'
             +' AND E_EXERCICE="' + QuelExo(DateToStr(Q.FindField('E_DATECOMPTABLE').AsDateTime)) + '"'
             +' AND E_DATECOMPTABLE="' + USDATETIME(Q.FindField('E_DATECOMPTABLE').AsDateTime) + '"'
             +' AND E_NUMEROPIECE=' + Q.FindField('E_NUMEROPIECE').AsString
             +' AND E_NUMLIGNE=' + Q.FindField('E_NUMLIGNE').AsString
             +' AND E_QUALIFPIECE="N"'
             +' AND E_NUMECHE=' + Q.FindField('E_NUMECHE').AsString,True);
  if (Not Q1.EOF) then begin
    O := TOBM.Create(EcrGen,'',False);
    O.ChargeMvt(Q1);
  end;
  Ferme(Q1);
  Result := O;
end;

procedure TOF_CPCFONBMP.InitTitres;
begin
  Case smp Of
     smpEnctraEnc : Ecran.Caption := HM.Mess[14]; // Export CFONB des traites-BOR à l''encaissement
     smpEncTraEsc : Ecran.Caption := HM.Mess[13]; // Export CFONB des traites-BOR à l''escompte
     smpEncPreBqe : Ecran.Caption := HM.Mess[15]; // Export CFONB des prélèvements
     smpDecVirbqe : Ecran.Caption := HM.Mess[16]; // Export CFONB des virements
   smpDecVirInbqe : Ecran.Caption := HM.Mess[18]; // Export CFONB des virements internationaux
     smpDecBorDec : Ecran.Caption := HM.Mess[17]; // Export CFONB des BOR
  end;
end;

procedure TOF_CPCFONBMP.Auxiliaire1Change(Sender: TObject);
begin
  SetControlText('E_AUXILIAIRE', GetControlText('AUXILIAIRE1'));
end;

procedure TOF_CPCFONBMP.Auxiliaire2Change(Sender: TObject);
begin
  SetControlText('E_AUXILIAIRE_', GetControlText('AUXILIAIRE2'));
end;

{***********A.G.L.***********************************************
Auteur  ...... : YMO
Créé le ...... : 12/04/2007
Modifié le ... :   /  /
Description .. : Branchement de la fiche auxiliaire
Mots clefs ... :
*****************************************************************}
procedure TOF_CPCFONBMP.AuxiElipsisClick( Sender : TObject );
begin
     THEdit(Sender).text:= CPLanceFiche_MULTiers('M;' +THEdit(Sender).text + ';' +THEdit(Sender).Plus + ';');
end;

procedure TOF_CPCFONBMP.BCtrlRibClick(Sender: TObject);
var
  StWRib : String ;
  i : Integer;
begin
  inherited;
  StWRib := RecupWhereCritere(TFMul(Ecran).Pages) ;
  if (StWRib = '') then Exit;

  // Si on n'est pas en Sélection inversée
  if ((Not FListe.AllSelected) and (FListe.NbSelected>0) and (FListe.NbSelected<100)) then begin
    // Si on n'a pas tous sélectionné ET qu'il y a au moins 1 et 100 au plus lignes sélectionnées
    // Rajoute une clause au WHERE
    StWRib := StWRib+' AND (';
    for i:=0 to FListe.NbSelected-1 do begin
      FListe.GotoLeBookmark(i) ;
{$IFDEF EAGLCLIENT}
      Q.TQ.Seek(FListe.Row-1);
{$ENDIF}
      StWRib := StWRib +' (E_NUMEROPIECE='+ Q.FindField('E_NUMEROPIECE').AsString +' AND E_NUMLIGNE='+ Q.FindField('E_NUMLIGNE').AsString +' AND E_JOURNAL="'+ Q.FindField('E_JOURNAL').AsString +'") OR';
    end;
    // Efface le dernier OR et rajoute ')'
    System.Delete(StWRib,length(StWRib)-2,3);
    StWRib := StWRib +')';
  end;
  If StWRib<>'' Then CPLanceFiche_VerifRib('WHERE='+StWRib);
end;

procedure TOF_CPCFONBMP.BOuvrirClick(Sender: TObject);
begin
  if Not ExporteSelection then Exit ;
  if Not FListe.AllSelected then FListe.ClearSelected
                            else FListe.AllSelected := False;
end;

procedure TOF_CPCFONBMP.bRibClick(Sender: TObject);
var
  O  : TOBM ;
  IsAux : Boolean ;
begin
  if Q.EOF then Exit ;
  O := GetLeOBM;
  IsAux := O.GetMvt('E_AUXILIAIRE')<>'';

  if O<>Nil then begin
    if (smp in [smpDecVirInBqe]) then begin
      if ModifRibOBM(O, True, False, '', IsAux, True) then BChercheClick(nil);
      end
    else begin
      if ModifRibOBM(O,True,FALSE,'',IsAux) then BChercheClick(nil);
    end;
  end;
  O.Free ;
end;

procedure TOF_CPCFONBMP.cExportClick(Sender: TObject);
begin
  case TCheckBox(GetControl('CEXPORT', True)).State of
    cbGrayed    : SetControlText('XX_WHEREEXPORT', '');
    cbChecked   : SetControlText('XX_WHEREEXPORT', 'E_CFONBOK="X"');
    cbUnchecked : SetControlText('XX_WHEREEXPORT', 'E_CFONBOK<>"X"');
  end;
end;

procedure TOF_CPCFONBMP.cRIBClick(Sender: TObject);
begin
  case TCheckBox(GetControl('CRIB', True)).State of
    cbGrayed    : SetControlText('XX_WHERERIB', '');
    cbChecked   : SetControlText('XX_WHERERIB', 'E_RIB<>""');
    cbUnchecked : SetControlText('XX_WHERERIB', 'E_RIB="" or E_RIB="////"');
  end;
end;

procedure TOF_CPCFONBMP.CTIDTICClick(Sender: TObject);
var
  b : Boolean;
begin
  b := not TCheckBox(GetControl('CTIDTIC')).Checked;
  SetControlVisible('AUXILIAIRE1', b);
  SetControlVisible('AUXILIAIRE2', b);
  SetControlVisible('HAUXILIAIRE1', b);
  SetControlVisible('HAUXILIAIRE2', b);
  if not b then begin
    SetControlText('AUXILIAIRE1', '');
    SetControlText('AUXILIAIRE2', '');
    If IsCli then SetControlProperty('E_GENERAL', 'DATATYPE', 'tzGTID')
             else SetControlProperty('E_GENERAL', 'DATATYPE', 'tzGTIC');
    SetControlText('E_GENERAL', '');
    SetControlText('XX_WHEREAUX', 'E_AUXILIAIRE="" AND E_NUMECHE>0 AND E_ETATLETTRAGE<>"RI" ');
    end
  else begin
    if IsCli then SetControlProperty('E_GENERAL', 'DATATYPE', 'tzGCollClient')
             else SetControlProperty('E_GENERAL', 'DATATYPE', 'tzGCollFourn');
    SetControlText('XX_WHEREAUX', 'E_AUXILIAIRE<>""');
  end;
end;

procedure TOF_CPCFONBMP.E_DATECOMPTABLEKeyPress(Sender: TObject; var Key: Char);
begin
  ParamDate(Ecran,Sender,Key);
end;

procedure TOF_CPCFONBMP.E_EXERCICEChange(Sender: TObject);
begin
  ExoToDates(GetControlText('E_EXERCICE'), GetControl('E_DATECOMPTABLE',True), GetControl('E_DATECOMPTABLE_',True) );
end;

procedure TOF_CPCFONBMP.FListeDblClick(Sender: TObject);
begin
  if Q.EOF then Exit ;
{$IFDEF EAGLCLIENT}
  Q.TQ.Seek(FListe.Row-1);
{$ENDIF}
  TrouveEtLanceSaisie(Q.TQ, taConsult, GetControlText('E_QUALIFPIECE'));
end;

procedure TOF_CPCFONBMP.BChercheClick(Sender: TObject);
begin
  TFMul(Ecran).BChercheClick(Sender); // inherited;
{$IFDEF EAGLCLIENT}
  if (Q.TQ <> nil) then SetControlEnabled('BRIB', not (Q.EOF));
{$ELSE}
  if (Q <> nil) then SetControlEnabled('BRIB', not (Q.EOF));
{$ENDIF}
end;

Initialization
  registerclasses ( [ TOF_CPCFONBMP ] ) ;
end.

