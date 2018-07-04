{***********UNITE*************************************************
Auteur  ...... : PH
Créé le ...... : 10/09/2001
Modifié le ... :   /  /
Description .. : Unit de gestion  des virements des acomptes
Mots clefs ... : PAIE;ACOMPTE
*****************************************************************
PT1    : 19/09/2001 SB V547  Emission du rib salarie erronée
                             Condition sur jointure et non sur WHERE
                             Fiche de bug N°310
PT2    : 28/11/2001 SB V563 Suit modif. AGL on ne redirige plus le bcherche
PT3    : 18/12/2001 SB V563 Nom du destinataire salarie erroné ds fichier
PT4    : 18/04/2002 SB V571 Fiche de bug n°10087 : V_PGI_env.LibDossier non
                            renseigné en Mono
PT5    : 13/05/2002 SB V571 Fiche de bug n°10099 suppression : pour traitement
                            internet
PT6    : 04/06/2002 PH V582 Gestion historique des évènements
PT7    : 20/09/2002 PH V585 Compatible eAGL
PT8    : 20/06/2005 PH V_60 FQ 12305 Génération du fichier de virements avec ou
                            sans séparateur
PT9    : 29/07/2005 PH V_60 FQ 12455 Erreur SQL ORACLE
PT10   : 16/06/2006 SB V_65 FQ 13091 Anomalie banque à émettre
PT11   : 14/03/2007 VG V_72 BQ_GENERAL n'est pas forcément unique
PT12   : 05/09/2007 FC V_80 FQ 14332 Salarié confidentiel visible
PT13   : 14/09/2007 PH V_80 FQ 14479 Prise en compte fourchette de date des acomptes et non le mois complet

}
unit UTOFPGVirementAcp;

interface
uses StdCtrls, Controls, Classes, sysutils, ComCtrls,
{$IFNDEF EAGLCLIENT}
  db, {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF} Mul,
  EdtREtat,
{$ELSE}
  UtileAgl, eMul,
{$ENDIF}
  HCtrls, HEnt1, HMsgBox, HQry, UTOF, UTOB, HTB97, ParamSoc,
  Dialogs,
  Windows,
  Commun,
  PgOutilsTreso, Constantes;

type
  TOF_PGVIREMENTACP = class(TOF)
    procedure OnArgument(stArgument: string); override;
    procedure OnLoad; override;
    procedure OnClose; override;
  private
    OkGeneration : Boolean;
    Chargement : Integer ; //PT-2 Chargement  PT10
    Separat: Boolean; // PT8
    DatePerEncours: TDateTime;
    procedure ChangePeriode(Sender: TObject);
    procedure AffectRib(Sender: TObject);
    procedure ReglementAcompte(Sender: TObject);
    procedure AffectDatePayeLe(PerEnCours: TdateTime);
    procedure LanceEtatAcp;
    function GenereFichier(Rib, RibMaj: string): Boolean;
    procedure DetruitFichier(NomFic: string);
    function RecupCritereAcp(Avec: boolean): string;
    procedure ExitEdit(Sender: TObject);
  end;

implementation
uses PGEditOutils2, P5Def, PGOperation, PGoutils, PgOutils2, EntPaie;
type
  Vir03 = record // ENREGISTREMENT DE TYPE ENTETE VIREMENT
    CodeEnr: string[2]; // Code enregistrement
    CodeOP: string[2]; // Code Operation
    ZR1: string[8]; // Zone reservee
    NumEmet: string[6]; // numéro emetteur
    ZR2: string[7]; // Zone reservee
    PayeLe: string[5]; // Date de paiement jjmma
    RaisonS: string[24]; // BQ_LIBELLE  raison sociale emetteur
    RefRem: string[7]; // reference de la remise
    ZR3: string[19]; // Zone reservee
    Monnaie: string[1]; // monnaie de la remise
    ZR4: string[5]; // Zone reservee
    CodeGui: string[5]; // BQ_GUICHET code quichet DO
    NumCpte: string[11]; // BQ_NUMEROCOMPTE numero de cpte DO
    ZR5: string[2]; // Zone reservee
    Identif: string[14]; // Identifiant DO
    ZR6: string[31]; // Zone reservée
    Banque: string[5]; // BQ_ETABBQ code banque du DO
    ZR7: string[6]; // Zone reservee
  end;
type
  Vir06 = record // ENREGISTREMENT DE TYPE SALARIE
    CodeEnr: string[2]; // Code enregistrement
    CodeOP: string[2]; // Code Operation
    ZR1: string[8]; // Zone reservee
    NumEmet: string[6]; // numéro emetteur
    RefInter: string[12]; // Reference interne
    NomDest: string[24]; // Nom du destinataire Nom Salarie
    Domicil: string[20]; // Domiciliation salarie R_DOMICILIATION
    NatEco: string[1]; // Nature Eco pour N.R
    Pays: string[3]; // Pays Eco pour N.R
    BalPay: string[8]; // Balance des paiements
    CodeGui: string[5]; // R_GUICHET code quichet salarie
    NumCpte: string[11]; // R_NUMEROCOMPTE numero de cpte salarie
    Montant: string[16]; // Montant net à payer PSD_MONTANT
    Libelle: string[29]; // Libellé := SALARIE : Numero du salarie
    CodeRej: string[2]; // Code rejet
    Banque: string[5]; // R_ETABBQ code banque
    ZR2: string[6]; // Zone reservee
  end;
type
  Vir08 = record // ENREGISTREMENT DE TYPE TOTAL DE LA REMISE
    CodeEnr: string[2]; // Code enregistrement
    CodeOP: string[2]; // Code Operation
    ZR1: string[8]; // Zone reservee
    NumEmet: string[6]; // numéro emetteur
    ZR2: string[84]; // Zone reservée
    Montant: string[16]; // Montant total de la remise
    ZR3: string[42]; // Zone reservee
  end;

  { TOF_PGVIREMENTACP }


procedure TOF_PGVIREMENTACP.OnArgument(stArgument: string);
var
  Ok: Boolean;
  CbxCodOp, CbxRibSal, cbMois, cbAnnee: THValComboBox;
  Min, Max, Mois, annee, Exer: string;
  Defaut: THEdit;
  DebExer, FinExer: TDateTime;
  Btn: TToolBarButton97;
  Check: TCheckBox;
begin
  inherited;
  Chargement := 0;//PT-2 PT10
  PGSuppression := 'ACP';
  //Evenements OnChange
//Btn:=TToolBarButton97(GetControl('BCherche')); PT-2
//if (Btn<>nil) then Btn.OnClick:=LanceCritere;  PT-2
  Btn := TToolBarButton97(GetControl('BOUVRIR'));
  if (Btn <> nil) then Btn.onclick := ReglementAcompte;
  CbxRibSal := THValComboBox(GetControl('PSA_RIBACPSOC'));
  if CbxRibSal <> nil then CbxRibSal.OnChange := AffectRib;

  // Affectation des valeurs par défaut;
  Defaut := ThEdit(getcontrol('DOSSIER'));
  if Defaut <> nil then
    //Defaut.text:=V_PGI_env.LibDossier;  //PT-4 Remplacé par GetParmsoc
    Defaut.text := GetParamSoc('SO_LIBELLE');
  Check := TCheckBox(GetControl('CKEURO'));
  if Check <> nil then Check.Checked := VH_Paie.PGTenueEuro;
  RecupMinMaxTablette('PG', 'SALARIES', 'PSA_SALARIE', Min, Max);
  Defaut := ThEdit(getcontrol('PSD_SALARIE'));
  if Defaut <> nil then
  begin
    Defaut.text := Min;
    Defaut.OnExit := ExitEdit;
  end;
  Defaut := ThEdit(getcontrol('PSD_SALARIE_'));
  if Defaut <> nil then
  begin
    Defaut.text := Max;
    Defaut.OnExit := ExitEdit;
  end;
  RecupMinMaxTablette('PG', 'ETABLISS', 'ET_ETABLISSEMENT', Min, Max);
  Defaut := ThEdit(getcontrol('PSD_ETABLISSEMENT'));
  if Defaut <> nil then Defaut.text := Min;
  Defaut := ThEdit(getcontrol('PSD_ETABLISSEMENT_'));
  if Defaut <> nil then Defaut.text := Max;
//  SetControlProperty('XX_WHERE', 'Text', ' PSA_RIBACPSOC="@@@" '); // PT9
  SetControlProperty('VCBXSUPPORT', 'value', 'DIS');
  CbxCodOp := THValComboBox(GetControl('VCBXCODEOP'));
  if CbxCodOp <> nil then CbxCodOp.value := '02';
  SetControlChecked('PSD_TOPREGLE', False);

  ok := RendExerSocialEnCours(Mois, annee, Exer, DebExer, FinExer);
  if Ok = True then
  begin
    DatePerEncours := EncodeDate(StrToInt(Annee), StrToInt(Mois), 1);
    AffectDatePayeLe(DatePerEncours);
    cbMois := THValComboBox(GetControl('CBMOIS'));
    if cbMois <> nil then
    begin
      if Length(Mois) = 1 then Mois := '0' + Mois;
      cbMois.value := Mois;
      CbMois.onchange := ChangePeriode;
    end;
    cbAnnee := THValComboBox(GetControl('CBANNEE'));
    if cbAnnee <> nil then
    begin
      CbAnnee.value := Exer;
      CbAnnee.onchange := ChangePeriode;
    end;
    Defaut := THEdit(GetControl('ECHEANCE')); //Date de paiement
    if Defaut <> nil then Defaut.text := GetControlText('PSD_DATEFIN');
  end;

//PT11
SetPlusBanqueCP (GetControl ('PSA_RIBACPSOC'));
SetPlusBanqueCP (GetControl ('VCBXRIBSAL'));
//FIN PT11
end;

procedure TOF_PGVIREMENTACP.OnClose;
begin
  inherited;

end;

procedure TOF_PGVIREMENTACP.OnLoad;
var
  Conf : String;
begin
  inherited;
  //DEB PT12
  Conf := '';
  if (V_PGI.Confidentiel = '0') then
  begin
    Conf := ' (PSD_CONFIDENTIEL="0") ';
    SetControlProperty('XX_WHERE', 'Text', Conf);
  end;
  //FIN PT12

  //DEB PT-2  PT10
  if Chargement < 2 then
   begin
    Chargement := Chargement + 1;
    exit;
  end; // FIN PT10
  if GetControlText('PSA_RIBACPSOC') = '' then
  begin
    PgiBox('Vous devez renseigner la banque du donneur d''ordre!', 'Virement des acomptes');
    exit;
  end;
  //FIN PT-2

  if (GetControlText('PSA_RIBACPSOC') <> '') then
    SetControlProperty('XX_WHERE', 'Text', Conf)
  else
    if Conf <> '' then
      SetControlProperty('XX_WHERE', 'Text', ' PSA_RIBACPSOC="@@@" AND ' + Conf)
    else
      SetControlProperty('XX_WHERE', 'Text', ' PSA_RIBACPSOC="@@@" ');
end;


procedure TOF_PGVIREMENTACP.AffectDatePayeLe(PerEnCours: TdateTime);
var
  Q: TQuery;
begin
  Q := OpenSql('SELECT MIN(PSD_DATEDEBUT) AS DEB,MAX(PSD_DATEFIN) AS FIN ' +
    'FROM HISTOSAISRUB ' +
    'WHERE PSD_DATEDEBUT>="' + USDateTime(PerEnCours) + '" ' +
    'AND PSD_DATEFIN<="' + USDateTime(FindeMois(PerEnCours)) + '" ' +
    'AND PSD_ORIGINEMVT="ACP" ', True);
  if not Q.eof then //PORTAGECWAS
  begin
    SetControlText('PSD_DATEDEBUT', DateToStr(Q.FindField('DEB').AsDateTime));
    SetControlText('PSD_DATEFIN', DateToStr(Q.FindField('FIN').AsDateTime));
    SetControlText('ECHEANCE', DateToStr(Q.FindField('FIN').AsDateTime));
  end
  else
  begin
    SetControlText('PSD_DATEDEBUT', '');
    SetControlText('PSD_DATEFIN', '');
    SetControlText('ECHEANCE', '');
  end;
  Ferme(Q);
end;

procedure TOF_PGVIREMENTACP.AffectRib(Sender: TObject);
var
  Vcbx, VcbxRib: ThValComboBox;
begin
  Vcbx := ThValComboBox(GetControl('PSA_RIBACPSOC'));
  VcbxRib := ThValComboBox(GetControl('VCBXRIBSAL'));
  if (Vcbx <> nil) and (VcbxRib <> nil) then
    VcbxRib.Value := Vcbx.Value;
end;

procedure TOF_PGVIREMENTACP.ChangePeriode(Sender: TObject);
var
  cbMois, cbAnnee: THValComboBox;
  Mois: WORD;
  Annee: string;
begin
  Mois := 1;
  cbMois := THValComboBox(GetControl('CBMOIS'));
  cbAnnee := THValComboBox(GetControl('CBANNEE'));
  if cbMois.Value <> '' then Mois := Trunc(StrToInt(cbMois.Value));
  if (cbMois.Value <> '') and (cbAnnee.Value <> '') then
  begin
    ControlMoisAnneeExer(cbMois.value, RechDom('PGANNEESOCIALE', cbAnnee.Value, FALSE), Annee);
    DatePerEncours := EncodeDate(StrToInt(Annee), Mois, 1);
    AffectDatePayeLe(DatePerEncours);
  end;
end;

function TOF_PGVIREMENTACP.GenereFichier(Rib, RibMaj: string): Boolean;
var
  F: TextFile;
  st, StPlus, FileN, BQNEmet, BQCGui, BQNCompte, BQCodeBq, LaDate: string;
  Salarie, Auxi: string;
  Reponse: WORD;
  V6: Vir06;
  V3: Vir03;
  V8: Vir08;
  MontantRemis: Double;
  VCBXOP, VCBXBQ, VCBXSupport: THValComboBox;
  TRIB_Sal, TR: TOB;
  ECH: THEdit;
  Q: TQuery;
  FichierBanq: string;
begin
if (VH_Paie.PgSeriaPaie = False) and (V_PGI.Debug = False) then
   begin
   PgiBox ('Vous devez sérialiser votre produit pour pouvoir générer un fichier!',
           'Sérialisation');
   result:= False;
   Exit;
   end;

if (GetControlText('CKSEPARAT')='X') then
   Separat:= TRUE
  else Separat := FALSE;
  // FIN PT8
  MontantRemis := 0;
  Result := False;
  VCBXOP := THValComboBox(GetControl('VCBXCODEOP'));
  if VCBXOP = nil then exit;
  VCBXBQ := THValComboBox(GetControl('VCBXRIBSAL'));  { PT10 PSA_RIBACPSOC }
  if VCBXBQ = nil then exit;
  ECH := THEdit(GetControl('ECHEANCE')); //Date de paiement
  if ECH = nil then exit;
  VCBXSupport := THValComboBox(GetControl('VCBXSUPPORT'));
  if VCBXSupport = nil then exit
  else
    if VCBXSupport.value = '' then
    begin
      PGIBox('Veuiller renseigner le type de support!', 'Virement des acomptes');
      exit;
    end;
  StPlus:= PGBanqueCP (True);     //PT11
  Q := OpenSql('SELECT BQ_REPVIR FROM BANQUECP WHERE BQ_GENERAL="' + RibMaj + '" '+StPlus, True);
  FichierBanq := Q.FindField('BQ_REPVIR').asstring;
  LaDate := DateToSTr(NOW);
  //Recupération du support de virement
  FileN := ConvertiFichierVirement(FichierBanq, VCBXSupport.Value);
  if FileExists(FileN) then
  begin
    reponse := HShowMessage('5;;Voulez-vous supprimer le fichier de virements ' + ExtractFileName(FileN) + ';Q;YN;Y;N', '', '');
    if reponse = 6 then DetruitFichier(FileN)
    else exit;
  end;
  // REcuperation des infos concernant la banque DO
  Q := OpenSQL('SELECT * FROM BANQUECP WHERE BQ_GENERAL="' + VCBXBQ.Value + '"'+StPlus, TRUE);
  if not Q.EOF then
  begin
    BQNEmet := Q.FindField('BQ_NUMEMETVIR').AsString;
    BQCGui := Q.FindField('BQ_GUICHET').AsString;
    BQNCompte := Q.FindField('BQ_NUMEROCOMPTE').AsString;
    BQCodeBq := Q.FindField('BQ_ETABBQ').AsString;
    Ferme(Q);
  end
  else
  begin
    Ferme(Q);
    exit;
  end;
  // Chargement des RIB SALARIES comprenant les rib pour les virements des acomptes
  TRIB_Sal := TOB.Create('Les RIB VIREMENTS SALARIES', nil, -1);
  Q := OpenSQL('SELECT R_AUXILIAIRE,R_DOMICILIATION,R_ETABBQ,R_NUMEROCOMPTE,R_GUICHET,' +
    'PSA_LIBELLE FROM RIB ' +
    'LEFT JOIN SALARIES ON PSA_AUXILIAIRE=R_AUXILIAIRE WHERE PSA_RIBACPSOC="' + rib + '"  ' + //PT-1
    'AND R_ACOMPTE="X"', TRUE);
  TRIB_Sal.LoadDetailDb('Les RIB salaries', '', '', Q, FALSE, FALSE);
  Ferme(Q);
  AssignFile(F, FileN);
{$I-}ReWrite(F);
{$I+}if IoResult <> 0 then
  begin
    PGIBox('Fichier inaccessible : ' + FileN, 'Abandon du traitement');
    Exit;
  end;
  // Entete du fichier identification du DO
  with V3 do
  begin
    CodeEnr := '03';
    CodeOP := Format_String(VCBXOP.value, 2);
    ZR1 := StringOfChar(' ', sizeof(ZR1)); // Zone reservee
    NumEmet := Format_String(BQNEmet, 6); // numéro emetteur
    ZR2 := StringOfChar(' ', sizeof(ZR2)); // Zone reservee
    PayeLe := Copy(ECH.Text, 1, 2) + Copy(ECH.Text, 4, 2) + Copy(ECH.Text, 10, 1); // Date de paiement jjmma
    RaisonS := Format_String(GetParamSoc('SO_LIBELLE'), 24); // BQ_LIBELLE  raison sociale emetteur
    RefRem := StringOfChar(' ', sizeof(RefRem)); // reference de la remise
    ZR3 := StringOfChar(' ', sizeof(ZR3)); // Zone reservee
    if VH_Paie.PGTenueEuro = True then Monnaie := 'E' // monnaie de la remise
    else if VH_Paie.PGMonnaieTenue = 'FRF' then Monnaie := 'F' else Monnaie := ' ';
    ZR4 := StringOfChar(' ', sizeof(ZR4)); // Zone reservee
    CodeGui := Format_String(BQCGui, 5); // BQ_GUICHET code quichet DO
    NumCpte := Format_String(BQNCompte, 11); // BQ_NUMEROCOMPTE numero de cpte DO
    ZR5 := StringOfChar(' ', sizeof(ZR5)); // Zone reservee
    Identif := StringOfChar(' ', sizeof(Identif)); // Identifiant DO
    ZR6 := StringOfChar(' ', sizeof(ZR6)); // Zone reservee
    Banque := Format_String(BQCodeBq, 11); // BQ_ETABBQ code banque du DO
    ZR7 := StringOfChar(' ', sizeof(ZR7)); // Zone reservee
  end;
  St := V3.CodeEnr + V3.CodeOP + V3.ZR1 + V3.NumEmet + V3.ZR2 + V3.PayeLe + V3.RaisonS + V3.RefRem
    + V3.ZR3 + V3.Monnaie + V3.ZR4 + V3.CodeGui + V3.NumCpte + V3.ZR5 + V3.Identif + V3.ZR6 + V3.Banque + V3.ZR7;
  if Separat then writeln(F, St)
  else write(F, St); // PT8
{$IFDEF EAGLCLIENT}
  TFMul(Ecran).Fetchlestous;
{$ENDIF}
  // Boucle sur la query = liste des virements à emettre = creation des enregistrements 06
  TFmul(Ecran).Q.First;
  while not TFmul(Ecran).Q.EOF do
  begin
    Salarie := TFmul(Ecran).Q.FindField('PSD_SALARIE').AsString;
    Auxi := TFmul(Ecran).Q.FindField('PSA_AUXILIAIRE').AsString;
    TR := TRIB_Sal.FindFirst(['R_AUXILIAIRE'], [Auxi], FALSE);
    if TR <> nil then
    begin
      with V6 do
      begin
        CodeEnr := '06';
        CodeOP := Format_String(VCBXOP.value, 2);
        ZR1 := StringOfChar(' ', sizeof(ZR1)); // Zone reservee
        NumEmet := Format_String(BQNEmet, 6); // numéro emetteur
        RefInter := StringOfChar(' ', sizeof(RefInter));
        NomDest := Format_String(TR.GetValue('PSA_LIBELLE'), 24); // Nom du destinataire Nom Salarie
        //PT-3 NomDest := Format_String (RechDom('TTBANQUECP',TFmul(Ecran).Q.FindField ('PSA_RIBACPSOC').AsString,False),24); // Nom du destinataire Nom Salarie
        Domicil := Format_String(TR.GetValue('R_DOMICILIATION'), 20); // Domiciliation salarie R_DOMICILIATION
        NatEco := StringOfChar(' ', sizeof(NatEco)); // Nature Eco pour N.R
        Pays := StringOfChar(' ', sizeof(Pays)); // Pays Eco pour N.R
        BalPay := StringOfChar(' ', sizeof(BalPay)); // Balance des paiements
        CodeGui := Format_String(TR.GetValue('R_GUICHET'), 5); // R_GUICHET code quichet salarie
        NumCpte := Format_String(TR.GetValue('R_NUMEROCOMPTE'), 11); // R_NUMEROCOMPTE numero de cpte salarie
        Montant := PGZeroAGauche(FloatToStr(Round(100.0 * TFmul(Ecran).Q.FindField('PSD_MONTANT').AsFloat)), 16); // Montant net à payer PVI_MONTANT
        // PT-5 13/05/2002 SB V571 Fiche de bug n°10099 suppression : pour traitement internet
        Libelle := Format_String('Salarie ' + TFmul(Ecran).Q.FindField('PSD_SALARIE').AsString, 29); // Libellé := SALARIE : Numero du salarie
        CodeRej := StringOfChar(' ', sizeof(CodeRej)); // Code rejet
        Banque := Format_String(TR.GetValue('R_ETABBQ'), 5); // R_GUICHET code quichet salarie
        ZR2 := StringOfChar(' ', sizeof(ZR2)); // Zone reservee
      end; // FIN with
      MontantRemis := MontantRemis + TFmul(Ecran).Q.FindField('PSD_MONTANT').AsFloat;
      St := V6.CodeEnr + V6.CodeOP + V6.ZR1 + V6.NumEmet + V6.RefInter + V6.NomDest + V6.Domicil + V6.NatEco
        + V6.Pays + V6.BalPay + V6.CodeGui + V6.NumCpte + V6.Montant + V6.Libelle + V6.CodeRej + V6.Banque + V6.ZR2;
      if Separat then writeln(F, St)
      else write(F, St); // PT8
    end // FIN if
    else
    begin //  salarie sans rib.
      PGIBox('Le salarié ' + Salarie + ' ne possède aucun rib!', 'Abandon du traitement');
      CloseFile(F);
      DetruitFichier(FileN);
      Exit;
    end;
    TFmul(Ecran).Q.Next;
  end; // FIN WHILE boucle sur la QUERY
  // Enregistrement 08 Total de la remise
  with V8 do
  begin
    CodeEnr := '08';
    CodeOP := Format_String(VCBXOP.value, 2);
    ZR1 := StringOfChar(' ', sizeof(ZR1)); // Zone reservee
    NumEmet := Format_String(BQNEmet, 6); // numéro emetteur
    ZR2 := StringOfChar(' ', sizeof(ZR2)); // Zone reservee
    Montant := PGZeroAGauche(FloatToStr(Round(100.0 * MontantRemis)), 16); // Montant total de la remise
    ZR3 := StringOfChar(' ', sizeof(ZR3)); // Zone reservee
  end;
  St := V8.CodeEnr + V8.CodeOP + V8.ZR1 + V8.NumEmet + V8.ZR2 + V8.Montant + V8.ZR3;
  if Separat then writeln(F, St)
  else write(F, St); // PT8
  CloseFile(F);
  if TRIB_Sal <> nil then TRIB_Sal.Free;
  PGIBox('Génération effectuée sous le fichier : ' + FileN + '', 'Génération des virements');
  result := True;
end;


procedure TOF_PGVIREMENTACP.DetruitFichier(NomFic: string);
begin
  DeleteFile(PChar(NomFic));
end;

procedure TOF_PGVIREMENTACP.ReglementAcompte(Sender: TObject);
var
  cbMois, cbAnnee, cbRib, cbRibMaj: THValComboBox;
  StOr, Mois, Annee, Rib, RibMaj, LibRibMaj, where: string;
  DateDeb: TDateTime;
  Btn: TToolBarButton97;
  Chk: TCheckBox;
  S1, S2: string;
  Trace: TStringList;
begin
  Btn := TToolBarButton97(GetControl('BCHERCHE'));
  cbMois := THValComboBox(GetControl('CBMOIS'));
  if cbMois = nil then exit;
  Mois := cbMois.value;
  cbAnnee := THValComboBox(GetControl('CBANNEE'));
  if cbAnnee = nil then exit;
  ControlMoisAnneeExer(Mois, RechDom('PGANNEESOCIALE', cbAnnee.Value, FALSE), Annee);
  DateDeb := EncodeDate(StrToInt(Annee), StrToInt(Mois), 1);
  cbRib := THValComboBox(GetControl('PSA_RIBACPSOC')); //DO initial
  cbRibMaj := THValComboBox(GetControl('VCBXRIBSAL')); //DO Mise à jour
  if cbRib <> nil then Rib := cbRib.value;
  if cbRibMAj <> nil then RibMaj := cbRibMaj.value;
{PT11
  LibRibMaj := RechDom('TTBANQUECP', RibMaj, False);
}
  LibRibMaj:= RechDom ('TTBANQUECP', RibMaj, False,
                       'BQ_NODOSSIER="'+V_PGI.NoDossier+'" AND'+BQCLAUSEWHERE);
//FIN PT11
  StOr := '';
  if Btn <> nil then Btn.Click;
  if (not TFmul(Ecran).Q.EOF) and (Rib <> '') then
  begin //Test si fichier imprimante
    if GetControlText('VCBXSUPPORT') = 'IMP' then
    begin
      LanceEtatAcp;
      Exit;
    end;
    OkGeneration := GenereFichier(Rib, RibMaj); //Génération du fichier
    if OkGeneration = True then
    begin //  MISE A JOUR de la table HISTOSAISRUB
      if MessageDlg('Génération du fichier effectuée.Voulez-vous une édition?', mtConfirmation, [mbYes, mbNo], 0) <> mrYes then
      else LanceEtatAcp;
      Where := RecupCritereAcp(False);
      ExecuteSQL('UPDATE HISTOSAISRUB SET PSD_TOPREGLE="X",PSD_RIBSALAIRE="' + RibMaj + '", ' +
        'PSD_BANQUEEMIS="' + LibRibMaj + '",PSD_DATEPAIEMENT="' + USDateTime(StrToDate(GetControlText('ECHEANCE'))) + '" ' +
        'WHERE PSD_ORIGINEMVT="ACP" ' + Where + '' +
        'AND PSD_DATEDEBUT>="' + USDateTime(StrToDate(GetControlText('PSD_DATEDEBUT'))) + '" ' + // PT13 USDateTime(DateDeb)
        'AND PSD_DATEFIN<="' + USDateTime(StrToDate(GetControlText('PSD_DATEFIN'))) + '" ');  // PT13 USDateTime(FindeMois(DateDeb))
      // PT6 04/06/2002 V582 PH Gestion historique des évènements
      Trace := TStringList.Create;
      S1 := DateToStr(DateDeb);
      S2 := DateToStr(FindeMois(DateDeb));
      Trace.add('Génération de ' + IntToStr(TFmul(Ecran).Q.RecordCount) + ' virements d'' acomptes du ' + S1 + ' au ' + S2);
      CreeJnalEvt('001', '009', 'OK', nil, nil, Trace);
      Trace.free;
      // FIN PT6
      if Btn <> nil then Btn.Click;
      // PT7  20/09/2002 V585 PH Compatible eAGL IsEmpty remplacé par EOF
      if TFmul(Ecran).Q.EOF then
      begin
        Chk := TCheckBox(GetControl('PSD_TOPREGLE'));
        if Chk <> nil then
          if Chk.State <> cbGrayed then Chk.State := cbGrayed;
        if Btn <> nil then Btn.Click;
      end;
    end
    else
    begin
      // PT6 04/06/2002 V582 PH Gestion historique des évènements
      Trace := TStringList.Create;
      S1 := DateToStr(DateDeb);
      S2 := DateToStr(FindeMois(DateDeb));
      Trace.add('Erreur génération des virements d''acomptes du ' + S1 + ' au ' + S2);
      CreeJnalEvt('001', '009', 'ERR', nil, nil, Trace);
      Trace.free;
      // FIN PT6
    end;
  end
  else
    if rib <> '' then
    begin
      if GetControlText('VCBXSUPPORT') = 'IMP' then
        PgiBox('Aucun enregistrement à imprimer!', 'Virement des acomptes')
      else PgiBox('Aucun enregistrement à générer!', 'Virement des acomptes');
    end;
end;

function TOF_PGVIREMENTACP.RecupCritereAcp(Avec: boolean): string;
var
  St, val: string;
  Chk: TCheckBox;
begin
  if Avec = True then St := 'WHERE PSD_SALARIE<>"00000000000" ' else St := '';
  val := GetControlText('PSD_PAYELE');
  if val <> '' then St := St + 'AND PSD_PAYELE>="' + USDateTime(StrToDate(val)) + '" ';
  val := GetControlText('PSD_PAYELE_');
  if val <> '' then St := St + 'AND PSD_PAYELE<="' + USDateTime(StrToDate(val)) + '" ';
  val := GetControlText('PSD_SALARIE');
  if val <> '' then St := St + 'AND PSD_SALARIE>="' + val + '" ';
  val := GetControlText('PSD_SALARIE_');
  if val <> '' then St := St + 'AND PSD_SALARIE<="' + val + '" ';
  val := GetControlText('PSD_ETABLISSEMENT');
  if val <> '' then St := St + 'AND PSD_ETABLISSEMENT>="' + val + '" ';
  val := GetControlText('PSD_ETABLISSEMENT_');
  if val <> '' then St := St + 'AND PSD_ETABLISSEMENT<="' + val + '" ';
  Chk := TCheckBox(GetControl('PSD_TOPREGLE'));
  if Chk <> nil then
  begin
    if Chk.State = cbGrayed then val := '';
    if Chk.State = cbChecked then val := 'X';
    if Chk.State = cbUnchecked then val := '-';
  end;
  if val = 'X' then St := St + 'AND PSD_TOPREGLE="' + val + '" ';
  if val = '-' then St := St + 'AND (PSD_TOPREGLE="' + val + '" OR PSD_TOPREGLE="") ';
  //PT12
  if (V_PGI.Confidentiel = '0') then
    St := St + 'AND (PSD_CONFIDENTIEL="0") ';
  result := St;
end;

procedure TOF_PGVIREMENTACP.LanceEtatAcp;
var
  Pages: TPageControl;
  StPlus, StSql, Temp: string;
begin
  temp := '';
  Pages := TPageControl(GetControl('Pages'));
  if pages = nil then exit;
  Temp := RecupWhereCritere(Pages);
  StPlus:= PGBanqueCP (True);      //PT11
  StSql := 'SELECT HISTOSAISRUB.*,PSA_LIBELLE,PSA_PRENOM,R_ETABBQ,R_NUMEROCOMPTE,' +
    'R_GUICHET,R_CLERIB,R_DOMICILIATION,BQ_LIBELLE,BQ_ETABBQ,BQ_CLERIB,BQ_GUICHET,' +
    'BQ_NUMEROCOMPTE FROM HISTOSAISRUB ' +
    'LEFT JOIN SALARIES ON PSA_SALARIE=PSD_SALARIE ' +
    'LEFT JOIN RIB ON R_AUXILIAIRE=PSA_AUXILIAIRE AND R_ACOMPTE="X" ' +
    'LEFT JOIN BANQUECP ON BQ_GENERAL="' + GetControlText('VCBXRIBSAL') + '" ' +
    '' + Temp + ' ' +
    'AND PSA_PAIACOMPTE="008"  AND BQ_GENERAL<>""'+StPlus+' AND PSD_MONTANT>0 ' + { DB2 SB 15/04/2005 }
    'ORDER BY PSD_DATEDEBUT,PSD_DATEFIN,PSD_RIBSALAIRE,PSD_SALARIE';

  LanceEtat('E', 'PRG', 'PAC', True, False, False, pages, StSql, '', False);
end;

procedure TOF_PGVIREMENTACP.ExitEdit(Sender: TObject);
var
  edit: thedit;
begin
  edit := THEdit(Sender);
  if edit <> nil then
    if (VH_Paie.PgTypeNumSal = 'NUM') and (length(Edit.text) < 11) and (isnumeric(edit.text)) then
      edit.text := AffectDefautCode(edit, 10);
end;

initialization
  registerclasses([TOF_PGVIREMENTACP]);
end.

