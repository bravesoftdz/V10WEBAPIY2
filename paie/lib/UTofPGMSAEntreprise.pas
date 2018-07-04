{***********UNITE*************************************************
Auteur  ...... : JL
Créé le ...... : 25/03/2004
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : MSA_ENTREPRISE ()
Mots clefs ... : TOF;MSA EDI
*****************************************************************
PT1 14/12/2004 JL V_60 FQ 11569 Affichage liste déclarants
PT2 17/02/2005 JL V_60 FQ 11569 Correction libellé déclarant
PT3 05/04/2005 JL V_60 Corrections format heures
PT4 19/05/2005 JL V_60 FQ 12292 Correction format total salair dans PE41
PT5 19/05/2005 JL V_60 FQ 12189 Alimentation TCP
PT6 18/10/2005 JL V_60 FQ 12189 Parametre société pour calcul TCP
                          + 23/11/2005 Ajout PHC_CUMULPAIE=35"
PT7 23/11/2005 JL V_65 FQ 12671 Gestion civilité
PT8 04/04/2006 JL V_65 Gestion UG MSA
PT9 17/05/2006 JL V_65 Gestion filtre sur établissement pour affichage grille dans MSA_ENTREPRISE
PT10 01/12/2006 JL V_750 Modifs pour nouveau cahier des charges, ajout nouvelles assiettes entreprise
PT11 02/02/2007 JL V_750 Modif tri enregistrement
PT12 18/07/2007 MF V_72 FQ 14588 : on encadre le nom du fichier (Chemin + Fichier) par des guillemets
PT13 08/08/2007 JL V_80 FQ 14477 Ajout saisie émetteur
PT14 13/11/2007 FC V_80 FQ 14892 Evolution cahier des charges Octobre 2007 + optimisation
PT15 22/01/2008 FC V_81 FQ 15124 Lignes PE41 à tort
PT16 24/01/2008 FC V_81 FQ 15176 L'UG (unité de gestion) n'est pas alimentée pour le segment PE41
PT17 20/05/2008 FC V_90 FQ 15385 Prendre en compte Paie décalée quand vérification des dates saisies
PT18 27/08/2008 FC V_810 FQ 15682 Modification cahier des charges - Modif tri du fichier EDI généré
}
unit UTofPGMSAEntreprise;

interface

uses
  {$IFDEF VER150}
  Variants,
  {$ENDIF}
  StdCtrls, Controls, Classes, Vierge,
  {$IFNDEF EAGLCLIENT}
  db, {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF} EdtREtat,
  {$ELSE}
  UtileAGL,
  {$ENDIF}
  forms, sysutils, ComCtrls, HCtrls, HEnt1, HMsgBox, UTOF, PG_OutilsEnvoi, UTob, PGOutils, Pgoutils2,AGLInit,
  ed_tools, ShellAPI, windows, EntPaie, ParamSoc;

type
  TOF_PGMSAENTREPRISE = class(TOF)
    procedure OnArgument(S: string); override;
  private
    DateDebut, DateFin: TDateTime;
    FEcrt, FZ, FFinal, FLect: textFile;
    GRem: THGrid;
    FiltreEtab : String;
    procedure RemplirGrilleRem;
    procedure ChargeZones();
    procedure Generation(Sender: TObject);
    procedure GenereFichier();
    procedure EcritureLigne(Valeur, Format: string; Longueur: integer);
    procedure EcritureLignePE11(Valeur, Format: string; Longueur: integer);
    function ControleFichierMSA(Valeur: string; NumSegment: Integer): string;
    function ControlePonctuation(var Donnee: string): string;
    procedure RecupFichierTemp;
    procedure AffichDeclarant(Sender:TObject);
  end;

implementation

procedure TOF_PGMSAENTREPRISE.OnArgument(S: string);
var
  Arg,St : string;
  QAttes : TQuery;
  Edit : THEdit;
begin
  inherited;
  Arg := S;
  St := 'PDA_ETABLISSEMENT = "" ' +
        ' AND (PDA_TYPEATTEST = "" OR PDA_TYPEATTEST LIKE "%MSA%" )  '; //PT1
  SetControlProperty('DECLARANT', 'Plus', St);
  QAttes := OpenSql('SELECT PDA_DECLARANTATTES FROM DECLARANTATTEST ' +
      'WHERE (PDA_TYPEATTEST = "" OR  PDA_TYPEATTEST LIKE "%MSA%" )  ' +
      'ORDER BY PDA_ETABLISSEMENT DESC', True);
  if not QAttes.eof then
  begin
        SetControlText('DECLARANT', QAttes.FindField('PDA_DECLARANTATTES').AsString);
        AffichDeclarant(nil);
  end;
  Edit := THEdit(GetControl('DECLARANT'));
  If Edit<>Nil Then Edit.OnExit := AffichDeclarant;
  dateDebut := StrToDate(Trim(ReadTokenPipe(Arg, ';')));
  DateFin := StrToDate(Trim(ReadTokenPipe(Arg, ';')));
  FiltreEtab := Trim(ReadTokenPipe(Arg, ';')); //PT9
  TFVierge(Ecran).BValider.OnClick := Generation;
  GRem := THGrid(GetControl('GREM'));
  if GRem <> nil then
  begin
    GRem.ColAligns[3] := taRightJustify;
    GRem.ColAligns[4] := taRightJustify;
    GRem.ColAligns[5] := taRightJustify;
    GRem.ColAligns[6] := taRightJustify;
    GRem.ColAligns[7] := taRightJustify;
    GRem.ColAligns[8] := taRightJustify;
    GRem.ColAligns[9] := taRightJustify;
    GRem.ColAligns[10] := taRightJustify;
    GRem.ColAligns[11] := taRightJustify;
    GRem.ColAligns[12] := taRightJustify;
    GRem.ColAligns[13] := taRightJustify;
    GRem.ColAligns[14] := taRightJustify;
    GRem.HideSelectedWhenInactive := true;
    GRem.ColWidths[0] := 125;
    GRem.ColWidths[1] := 125;
    GRem.ColWidths[2] := 60;
    GRem.ColWidths[3] := 80;
    GRem.ColWidths[4] := 80;
    GRem.ColWidths[5] := 80;
    GRem.ColWidths[6] := 80;
    GRem.ColWidths[7] := 80;
    GRem.ColWidths[8] := 80;
    GRem.ColWidths[9] := 80;
    GRem.ColWidths[10] := 80;
    GRem.ColWidths[11] := 80;
    GRem.ColWidths[12] := 80;           //PT18
    GRem.ColWidths[13] := 80;           //PT18
    GRem.CellValues[4, 1] := 'cas standard';
    GRem.CellValues[5, 1] := 'rempl. non imposés';
    GRem.CellValues[6, 1] := 'rempl. imposés 6,2';
    GRem.CellValues[7, 1] := 'rempl. imposés 6,6';
    GRem.CellValues[8, 1] := '66';           //PT18
    GRem.CellValues[9, 1] := '67';           //PT18
    GRem.CellValues[10, 1] := '68';
    GRem.CellValues[11, 1] := '58';
    GRem.CellValues[12, 1] := '59';
    GRem.CellValues[13, 1] := '61';
    GRem.CellValues[14, 1] := '62';
    GRem.ColFormats[0] := 'CB=TTETABLISSEMENT';
    GRem.ColFormats[1] := 'CB=PGMSAACTIVITE';
    GRem.ColFormats[2] := '';
    GRem.ColFormats[3] := '# ##0.00';
    GRem.ColFormats[4] := '# ##0.00';
    GRem.ColFormats[5] := '# ##0.00';
    GRem.ColFormats[6] := '# ##0.00';
    GRem.ColFormats[7] := '# ##0.00';
    GRem.ColFormats[8] := '# ##0.00';
    GRem.ColFormats[9] := '# ##0.00';
    GRem.ColFormats[10] := '# ##0.00';
    GRem.ColFormats[11] := '# ##0.00';
    GRem.ColFormats[12] := '# ##0.00';
    GRem.ColFormats[13] := '# ##0.00'; //PT18
    GRem.ColFormats[14] := '# ##0.00'; //PT18
  end;
  ChargeZones();
  //DEBUT PT13
  SetControlText('EMETSOC2', GetParamSocSecur('SO_PGEMETTEUR',''));
  SetControlText ('LEMETSOC2', RechDom ('PGEMETTEURSOC',
                                         GetControlText ('EMETSOC2'), False));
   //FIN PT13
end;

procedure TOF_PGMSAENTREPRISE.RemplirGrilleRem;
var
  Q: TQuery;
  TobEtab, TobGrille, TobActivite, TG,TobUG: Tob;
  WhereEtab,Etab, Activite,UGMsa : string;
  i, a,u: Integer;
  Montant: Double;
  CalculerTcp : Boolean;
  St : String;
begin
  CalculerTcp := GetParamSocSecur ('SO_PGMSACALCTCP',False);
  WhereEtab := '';
  If FiltreEtab <> '' then WhereEtab := ' WHERE ET_ETABLISSEMENT="'+FiltreEtab+'"'; //PT9
  Q := OpenSQL('SELECT ET_ETABLISSEMENT FROM ETABLISS'+WhereEtab, True);
  TobEtab := Tob.Create('LesEtablissements', nil, -1);
  TobEtab.LoadDetailDB('LesEtablissements', '', '', Q, False);
  Ferme(Q);
  TobGrille := Tob.Create('RemplirLaGrille', nil, -1);
  for i := 0 to TobEtab.Detail.Count - 1 do
  begin
    Etab := TobEtab.Detail[i].GetValue('ET_ETABLISSEMENT');
//Optimisation     Q := OpenSQL('SELECT DISTINCT (PE3_MSAACTIVITE) FROM MSAPERIODESPE31 WHERE PE3_ETABLISSEMENT="' + Etab + '"' +
//      ' AND PE3_DATEDEBUT>="' + UsDatetime(Datedebut) + '" AND PE3_DATEFIN<="' + UsDatetime(DateFin) + '"', True);
//    TobActivite.LoadDetailDB('LesActivites', '', '', Q, False);
    St := 'SELECT DISTINCT (PE3_MSAACTIVITE) FROM MSAPERIODESPE31 WHERE PE3_ETABLISSEMENT="' + Etab + '"' +
      ' AND PE3_DATEDEBUT>="' + UsDatetime(Datedebut) + '" AND PE3_DATEFIN<="' + UsDatetime(DateFin) + '"';
    TobActivite := Tob.Create('LesActivites', nil, -1);
    TobActivite.LoadDetailDbFromSQL('LesActivites', St);
//    Ferme(Q);
    for a := 0 to TobActivite.Detail.Count - 1 do
    begin
      Activite := TobActivite.detail[a].GetValue('PE3_MSAACTIVITE');
//Optimisation      Q := OpenSQL('SELECT DISTINCT (PE3_UGMSA) FROM MSAPERIODESPE31 WHERE PE3_MSAACTIVITE="'+Activite+'" AND PE3_ETABLISSEMENT="' + Etab + '"' +
//      ' AND PE3_DATEDEBUT>="' + UsDatetime(Datedebut) + '" AND PE3_DATEFIN<="' + UsDatetime(DateFin) + '"', True);
//      TobUG.LoadDetailDB('LesUG', '', '', Q, False);
      St := 'SELECT DISTINCT (PE3_UGMSA) FROM MSAPERIODESPE31 WHERE PE3_MSAACTIVITE="'+Activite+'" AND PE3_ETABLISSEMENT="' + Etab + '"' +
      ' AND PE3_DATEDEBUT>="' + UsDatetime(Datedebut) + '" AND PE3_DATEFIN<="' + UsDatetime(DateFin) + '"';
      TobUG := Tob.Create('LesUG', nil, -1);
      TobUG.LoadDetailDbFromSQL('LesUG', St);
//      Ferme(Q);
      for u := 0 to TobUG.Detail.Count - 1 do
      begin
           UGMsa := TobUG.detail[u].GetValue('PE3_UGMSA');
            Q := OpenSQL(' SELECT SUM (PHB_BASECOT) AS TOTAL FROM' +
            ' HISTOBULLETIN' +
            ' LEFT JOIN COTISATION ON ##PCT_PREDEFINI## PHB_RUBRIQUE=PCT_RUBRIQUE AND PHB_NATURERUB=PCT_NATURERUB' +
            ' LEFT JOIN MSAPERIODESPE31 ON PHB_SALARIE=PE3_SALARIE AND PHB_DATEDEBUT=PE3_DATEDEBUT AND PHB_DATEFIN=PE3_DATEFIN AND PE3_ORDRE=0' +
            ' WHERE PHB_DATEFIN>="' + UsDateTime(DateDebut) + '" AND PHB_DATEFIN<="' + UsDateTime(DateFin) + '"' +
            ' AND PCT_BASECSGCRDS="X" AND PE3_MSAACTIVITE="' + Activite + '" AND PE3_ETABLISSEMENT="' + Etab + '" AND PE3_UGMSA="' + UGMsa + '"', True);
            Montant := 0;
            if not Q.eof then Montant := Q.Findfield('TOTAL').AsFloat; // PortageCWAS
            Ferme(Q);
            TG := Tob.Create('UneFille', TobGrille, -1);
            TG.AddChampSupValeur('ETABLISSEMENT', Etab, False);
            TG.AddChampSupValeur('ACTIVITE', Activite, False);
            TG.AddChampSupValeur('UG', UGMsa, False);
            TG.AddChampSupValeur('CSG', Montant, False);
            //DEBUT PT5
            If CalculerTcp then
            begin
                  if (VH_PAIE.PGProfilFnal <> '') and (VH_PAIE.PGProfilFnal <> NULL) then
                  begin
                       Q := OpenSQL('SELECT SUM(PHC_MONTANT) AS MONTANT ' +
                       'FROM HISTOCUMSAL '+
                       ' LEFT JOIN MSAPERIODESPE31 ON PHC_SALARIE=PE3_SALARIE AND PHC_DATEDEBUT=PE3_DATEDEBUT AND PHC_DATEFIN=PE3_DATEFIN AND PE3_ORDRE=0' +
                       'WHERE PHC_CUMULPAIE="35" AND PHC_DATEDEBUT>="' + UsDateTime(DateDebut) + '"  ' +
                       'AND PHC_DATEFIN<="' + UsDateTime(DateFin) + '" AND PE3_ETABLISSEMENT="' + Etab + '"',true);
                       If Not Q.Eof then TG.AddChampSupValeur('TCP', Q.FindField('MONTANT').AsFloat, False)
                       else TG.AddChampSupValeur('TCP', 0, False);
                       Ferme(Q);
                  end
                  else TG.AddChampSupValeur('TCP', 0, False);
            end
            else TG.AddChampSupValeur('TCP', 0, False);
            //FIN PT5
          end;
          FreeAndNil(TobUG);
      end;
      FreeAndNil(TobActivite);
  end;
  FreeAndNil(TobEtab);
  GRem.Rowcount := 2 + TobGrille.Detail.Count;
  for i := 0 to TobGrille.Detail.Count - 1 do
  begin
    GRem.CellValues[0, i + 2] := TobGrille.Detail[i].GetValue('ETABLISSEMENT');
    GRem.CellValues[1, i + 2] := TobGrille.Detail[i].GetValue('ACTIVITE');
    GRem.CellValues[2, i + 2] := TobGrille.Detail[i].GetValue('UG');
    GRem.CellValues[4, i + 2] := FloatToStr(Arrondi(TobGrille.Detail[i].GetValue('CSG'), 0));
    GRem.CellValues[3, i + 2] := FloatToStr(Arrondi(TobGrille.Detail[i].GetValue('TCP'), 0)); //PT5
  end;
  FreeAndNil(TobGrille);
end;

procedure TOF_PGMSAENTREPRISE.ChargeZones();
var
  Siren, Siret: string;
begin
  SetControlText('DATEDEBUT', DateToStr(DateDebut));
  SetControlText('DATEFIN', DateToStr(DateFin));
  SetControlText('PERIODICITE', 'T');
  SetControlText('ERAISONSOC', GetParamSoc('SO_LIBELLE'));
  Siret := GetParamSoc('SO_SIRET');
  ForceNumerique(Siret, Siren);
  SetControlText('ESIREN', Copy(Siren, 1, 9));
  RemplirGrilleRem;
end;

procedure TOF_PGMSAENTREPRISE.Generation(Sender: TObject);
begin
  try
    begintrans;
    GenereFichier();
    CommitTrans;
  except
    Rollback;
    PGIBox('Une erreur est survenue lors de la mise à jour de la base', Ecran.Caption);
  end;
end;

procedure TOF_PGMSAENTREPRISE.GenereFichier();
var
  EnregEnvoi: TEnvoiSocial;
  Taille: longint;
  QRechDelete: TQuery;
  StDelete, fichier, Libelle, Trimestre, TrimestreDeb, TrimestreFin, Annee: string;
  Ordre, i, j, e, a,u,elt,r,k: Integer;
  Q: TQuery;
  FileN, FileZ, FileTemp: string;
  reponse, rep: integer;
  TabMess: array[1..22] of string;
  MontantSalaires, MontantSalairesEtab: Double;
  NbMess, NumElement, NbEnreg, NbSalaries: Integer;
  Erreur, DeclarerAsiette: Boolean;
  TobMSA, TobMSA2, TobEvolutions, TobEtablissement, TobActivite, TA, TobLesAssiettes, TLA,TobUG: Tob;
  MessError, StPages: string;
  IdEtab, CodeActivite, CodeActSal, UG, CodeArt, IdSal, Nom, Prenom, TopNdP, PourcExo, NbJ, NbH, EltC1, EltC2, EltC3, EltC4, EltC5: string;
  EltC6, EltC7, EltC8, EltC9, EltC10,EltC11,EltC12,Velt6, Velt7, Velt8, Velt9, Velt10,Velt11,Velt12 : String;
  Velt1, Velt2, Velt3, Velt4, Velt5, SDateNaiss, SDateDebSal, SDateFinSal: string;
  DateNaiss, DateDebSal, DateFinSal: TDateTime;
  IdEntreprise, RaisonSoc, TypeFic,UniteGestion: string;
  SDebAct, SFinAct, SDateDebut, SDateFin: string;
  DebAct, FinAct, DebSuspCt, FinSuspCt, DateEffet: TDateTime;
  SDebSuspCt, SFinSuspCt, CodeSuspCt: string;
  SDateEffet, DepartLieuTrav, CommLieuTrav, PTpsPart, CodeTpsPart, CodeCDD, NbHConvention, Niveau: string;
  Etablissement, ActiviteMSA: string;
  PersTech, TopCadre, PolyEmp, PersTechCUMA, FiscHorsF, CodeClasse, Qualite, CRCCA, AUBRY1, AUBRY2, Emploi, Retraite, Convention: string;
  TotalSalaires: string;
  Pages: TPageControl;
  MontantSalairesUG : Double;
  CollegeP,SectionP : String;
  TabElt : Array [1..12] of string;
  TypeRem : Array[1..4] of String;
  TauxMajo : Array[1..4] of String;
  NbHeures : Array[1..4] of String;
  MontantRem : Array[1..4] of String;
  MemSalarie:String;
  MemDateDebut,MemDateFin : TDateTime;
  St : String;
  TrimestreDebExerc,TrimestreFinExerc : Array[1..8] of String;
  z:integer;
begin
  Pages := TPageControl(GetControl('PAGES'));
  IdEntreprise := GetControlText('ESIREN');
  if IDEntreprise = '' then
  begin
    PGIBox('Vous devez renseigner le n° SIREN', Ecran.Caption);
    Exit;
  end;
  if ControlSiret(IdEntreprise) = False then
  begin
    PGIBox('Le SIRET n''est pas valide', 'Saisie Entreprise MSA');
    Exit;
  end;
  Datedebut := StrToDate(GetControlText('DATEDEBUT'));
  DateFin := StrTodate(GetControlText('DATEFIN'));
  Annee := FormatDateTime('yyyy', DateDebut);
  SetControlText('ANNEE', Annee);
  Erreur := False;
  TrimestreDeb := FormatDateTime('ddmm', DateDebut);
  TrimestreFin := FormatDateTime('ddmm', DateFin);
  if TrimestreFin = '2902' then TrimestreFin := '2802';  //PT17
  SDateDebut := FormatDateTime('yyyymmdd', DateDebut);
  SDateFin := FormatDateTime('yyyymmdd', DateFin);
  //DEB PT17
  if GetParamSocSecur('SO_PGDECALAGE',False) then
  begin
    TrimestreDebExerc[1] := '0112';
    TrimestreFinExerc[1] := '2802';
    TrimestreDebExerc[2] := '0103';
    TrimestreFinExerc[2] := '3105';
    TrimestreDebExerc[3] := '0106';
    TrimestreFinExerc[3] := '3108';
    TrimestreDebExerc[4] := '0109';
    TrimestreFinExerc[4] := '3011';
  end
  else
  begin
    TrimestreDebExerc[1] := '0101';
    TrimestreFinExerc[1] := '3103';
    TrimestreDebExerc[2] := '0104';
    TrimestreFinExerc[2] := '3006';
    TrimestreDebExerc[3] := '0107';
    TrimestreFinExerc[3] := '3009';
    TrimestreDebExerc[4] := '0110';
    TrimestreFinExerc[4] := '3112';
  end;
  //FIN PT17
  if TrimestreDeb = TrimestreDebExerc[1] then //PT17 '0101' then
  begin
    if TrimestreFin <> TrimestreFinExerc[1] then Erreur := True //PT17 '3103' then Erreur := True
    else Trimestre := '1';
  end
  else
    if TrimestreDeb = TrimestreDebExerc[2] then //PT17 '0104' then
  begin
    if TrimestreFin <> TrimestreFinExerc[2] then Erreur := True //PT17 '3006' then Erreur := True
    else Trimestre := '2';
  end
  else
    if TrimestreDeb = TrimestreDebExerc[3] then //PT17 '0107' then
  begin
    if TrimestreFin <> TrimestreFinExerc[3] then Erreur := True //PT17 '3009' then Erreur := True
    else Trimestre := '3';
  end
  else
    if TrimestreDeb = TrimestreDebExerc[4] then //PT17 '0110' then
  begin
    if TrimestreFin <> TrimestreFinExerc[4] then Erreur := True //PT17 '3112' then Erreur := True
    else Trimestre := '4';
  end
  else Erreur := True;
  if Erreur = True then
  begin
    PGIBox('les dates saisies ne correspondent pas à un trimestre', Ecran.Caption);
    Exit;
  end;
  if Trimestre = '1' then SetControlText('NUMTRIM', Trimestre + ' er')
  else SetControlText('NUMTRIM', Trimestre + ' ème');
  Fichier := 'MSA_' + IDEntreprise + '_' + Annee + Trimestre + '.MSA';
  Libelle := 'Déclaration des salaires ' + getControlText('DATEDEBUT') + ' au ' + GetControlText('DATEFIN');
  IdEntreprise := GetControlText('ESIREN');
  {$IFNDEF EAGLCLIENT}
  FileN := V_PGI.DatPath + '\' + Fichier;
  FileTemp := V_PGI.DatPath + '\MSATEMP.MSA';
  {$ELSE}
  FileN := VH_Paie.PGCheminEagl + '\' + Fichier;
  FileTemp := VH_Paie.PGCheminEagl + '\MSATEMP.MSA';
  {$ENDIF}
  if FileExists(FileN) then
  begin
    reponse := HShowMessage('5;;Voulez-vous supprimer votre fichier MSA' + FileN + ';Q;YN;Y;N', '', '');
    if reponse = 6 then DeleteFile(PChar(FileN))
    else exit;
  end;
  ExecuteSQL('DELETE FROM ENVOISOCIAL WHERE PES_TYPEMESSAGE="MSA" AND '+  // Suppression des envois existants avec les mêmes dates
  'PES_DATEDEBUT="'+UsDateTime(Datedebut)+'" AND PES_DATEFIN="'+UsDateTime(DateFin)+'" AND PES_SIRETDO="'+GetControlText('ESIREN')+'"');
  if FileExists(FileTemp) then DeleteFile(PChar(FileTemp));
  AssignFile(FEcrt, FileTemp);
  AssignFile(FFinal, FileN);
  {$I-}
  ReWrite(FEcrt);
  {$I+}
  if IoResult <> 0 then
  begin
    PGIBox('Fichier inaccessible : ' + FileTemp, 'Abandon du traitement');
    Exit;
  end;
  {$I-}
  ReWrite(FFinal);
  {$I+}
  if IoResult <> 0 then
  begin
    PGIBox('Fichier inaccessible : ' + FileN, 'Abandon du traitement');
    Exit;
  end;
  for i := 1 to 22 do TabMess[i] := '';
  {$IFNDEF EAGLCLIENT}
  if V_PGI.DatPath <> '' then FileZ := V_PGI.DatPath + '\ErreurMSA.log'
  else FileZ := 'C:\ErreurMSA.log';
  {$ELSE}
  if VH_Paie.PGCheminEagl <> '' then FileZ := VH_Paie.PGCheminEagl + '\ErreurMSA.log'
  else FileZ := 'C:\ErreurMSA.log';
  {$ENDIF}
  if SupprimeFichier(FileZ) = False then exit;
  AssignFile(FZ, FileZ);
  {$I-}ReWrite(FZ);
  {$I+}
  if IoResult <> 0 then
  begin
    PGIBox('Fichier inaccessible : ' + FileZ, 'Abandon du traitement');
    Exit;
  end;
  writeln(FZ, 'Création fichier MSA : Gestion des messages d''erreur.');
  RaisonSoc := GetControlText('ERAISONSOC');
  TypeFic := '';
  NbEnreg := 0;
//Optimisation  Q := OpenSQL('SELECT ET_LIBELLE,ET_ETABLISSEMENT,ET_SIRET FROM ETABLISS', True);
//  TobEtablissement.LoadDetailDB('LesEtablissements', '', '', Q, False);
  St := 'SELECT ET_LIBELLE,ET_ETABLISSEMENT,ET_SIRET FROM ETABLISS';
  TobEtablissement := Tob.Create('LesEtablissement', nil, -1);
  TobEtablissement.LoadDetailDbFromSQL('LesEtablissements', St);
//  Ferme(Q);
  TobLesAssiettes := Tob.Create('LesAssiettes', nil, -1);
  for i := 2 to GRem.RowCount - 1 do
  begin
     begin
          TLA := Tob.Create('UneAssiette', TobLesAssiettes, -1);
          TLA.AddChampSupValeur('ETABLISSEMENT', Grem.CellValues[0, i], False);
          TLA.AddChampSupValeur('ACTIVITE', Grem.CellValues[1, i], False);
          TLA.AddChampSupValeur('UG', Grem.CellValues[2, i], False);
          if isNumeric(Grem.CellValues[3, i]) then TLA.AddChampSupValeur('51', StrToFloat(Grem.CellValues[3, i]), False)
          else TLA.AddChampSupValeur('51', 0, False);
          if isNumeric(Grem.CellValues[4, i]) then TLA.AddChampSupValeur('53', StrToFloat(Grem.CellValues[4, i]), False)
          else TLA.AddChampSupValeur('53', 0, False);
          if isNumeric(Grem.CellValues[5, i]) then TLA.AddChampSupValeur('54', StrToFloat(Grem.CellValues[5, i]), False)
          else TLA.AddChampSupValeur('54', 0, False);
          if isNumeric(Grem.CellValues[6, i]) then TLA.AddChampSupValeur('56', StrToFloat(Grem.CellValues[6, i]), False)
          else TLA.AddChampSupValeur('56', 0, False);
          if isNumeric(Grem.CellValues[7, i]) then TLA.AddChampSupValeur('63', StrToFloat(Grem.CellValues[7, i]), False)
          else TLA.AddChampSupValeur('63', 0, False);
          //DEB PT18
          if isNumeric(Grem.CellValues[8, i]) then TLA.AddChampSupValeur('66', StrToFloat(Grem.CellValues[8, i]), False)
          else TLA.AddChampSupValeur('66', 0, False);
          if isNumeric(Grem.CellValues[9, i]) then TLA.AddChampSupValeur('67', StrToFloat(Grem.CellValues[9, i]), False)
          else TLA.AddChampSupValeur('67', 0, False);
          if isNumeric(Grem.CellValues[10, i]) then TLA.AddChampSupValeur('68', StrToFloat(Grem.CellValues[10, i]), False)
          else TLA.AddChampSupValeur('68', 0, False);
          //FIN PT18
          if isNumeric(Grem.CellValues[11, i]) then TLA.AddChampSupValeur('58', StrToFloat(Grem.CellValues[11, i]), False)
          else TLA.AddChampSupValeur('58', 0, False);
          if isNumeric(Grem.CellValues[12, i]) then TLA.AddChampSupValeur('59', StrToFloat(Grem.CellValues[12, i]), False)
          else TLA.AddChampSupValeur('59', 0, False);
//PT18          if isNumeric(Grem.CellValues[10, i]) then TLA.AddChampSupValeur('60', StrToFloat(Grem.CellValues[10, i]), False)
//PT18          else TLA.AddChampSupValeur('60', 0, False);
          if isNumeric(Grem.CellValues[13, i]) then TLA.AddChampSupValeur('61', StrToFloat(Grem.CellValues[13, i]), False)
          else TLA.AddChampSupValeur('61', 0, False);
          if isNumeric(Grem.CellValues[14, i]) then TLA.AddChampSupValeur('62', StrToFloat(Grem.CellValues[14, i]), False)
          else TLA.AddChampSupValeur('62', 0, False);
     end;
  end;
//Optimisation  Q := OpenSQL('SELECT CC_ABREGE FROM CHOIXCOD WHERE CC_TYPE="PMS"', True);
//  TobActivite.LoadDetailDB('LesActivites', '', '', Q, False);
  St := 'SELECT CC_ABREGE FROM CHOIXCOD WHERE CC_TYPE="PMS"';
  TobActivite := Tob.Create('LesActivites', nil, -1);
  TobActivite.LoadDetailDbFromSQL('LesActivites', St);
//  Ferme(Q);
  NbSalaries := 0;
  MontantSalaires := 0;
  if TobActivite.Detail.Count = 0 then
  begin
    TA := Tob.Create('SansActivite', TobActivite, -1);
    TA.AddChampSupValeur('CC_ABREGE', '', False);
  end;
  try
    begintrans;
    InitMoveProgressForm(nil, 'Début du traitement', 'Veuillez patienter SVP ...', TobEtablissement.Detail.Count - 1, FALSE, TRUE);
    for e := 0 to TobEtablissement.Detail.Count - 1 do
    begin
      Etablissement := TobEtablissement.Detail[e].GetValue('ET_ETABLISSEMENT');
      IDEtab := TobEtablissement.Detail[e].GetValue('ET_SIRET');
      for a := 0 to TobActivite.Detail.Count - 1 do
      begin
        MontantSalairesEtab := 0;
        ActiviteMSA := TobActivite.Detail[a].GetValue('CC_ABREGE');
//Optimisation        Q := OpenSQL('SELECT DISTINCT(PE3_UGMSA) UG FROM MSAPERIODESPE31 WHERE' +
//               ' PE3_DATEDEBUT>="' + UsDatetime(Datedebut) + '"' +
//               ' AND PE3_DATEFIN<="' + UsDatetime(DateFin) + '"' +
//               ' AND PE3_MSAACTIVITE="' + ActiviteMSA + '"' +
//               ' AND PE3_ETABLISSEMENT="' + Etablissement + '"', True);
//        TobUG.LoadDetailDB('MSA', '', '', Q, False);
        St :='SELECT DISTINCT(PE3_UGMSA) UG FROM MSAPERIODESPE31 WHERE' +
               ' PE3_DATEDEBUT>="' + UsDatetime(Datedebut) + '"' +
               ' AND PE3_DATEFIN<="' + UsDatetime(DateFin) + '"' +
               ' AND PE3_MSAACTIVITE="' + ActiviteMSA + '"' +
               ' AND PE3_ETABLISSEMENT="' + Etablissement + '"';
        TobUG := Tob.create('MSA', nil, -1);
        TobUG.LoadDetailDbFromSQL('MSA', St);
        if TobUG.Detail.Count = 0 then DeclarerAsiette := False; //PT15
//        Ferme(Q);
               //SEGMENT PE21   PT11 Tri par enregistrement modifié
//Optimisation        Q := OpenSQL('SELECT MSAEVOLUTIONSPE2.*,PSA_LIBELLE,PSA_PRENOM,PSA_ETABLISSEMENT,PSA_NUMEROSS,PSA_DATENAISSANCE FROM MSAEVOLUTIONSPE2 ' +
//          'LEFT JOIN SALARIES ON PE2_SALARIE=PSA_SALARIE ' +
//          'WHERE PE2_TYPEEVOLMSA="PE21" ' +
//          'AND PE2_MSAACTIVITE="' + ActiviteMSA + '" AND PE2_ETABLISSEMENT="' + Etablissement + '" ' +
//          'AND PE2_DATEEFFET>="' + UsdateTime(DateDebut) + '" AND PE2_DATEEFFET<="' + UsDateTime(DateFin) + '"', True);
//        TobEvolutions.LoadDetailDB('LesEvolutions', '', '', Q, false);
        St := 'SELECT MSAEVOLUTIONSPE2.*,PSA_LIBELLE,PSA_PRENOM,PSA_ETABLISSEMENT,PSA_NUMEROSS,PSA_DATENAISSANCE FROM MSAEVOLUTIONSPE2 ' +
          'LEFT JOIN SALARIES ON PE2_SALARIE=PSA_SALARIE ' +
          'WHERE PE2_TYPEEVOLMSA="PE21" ' +
          'AND PE2_MSAACTIVITE="' + ActiviteMSA + '" AND PE2_ETABLISSEMENT="' + Etablissement + '" ' +
          'AND PE2_DATEEFFET>="' + UsdateTime(DateDebut) + '" AND PE2_DATEEFFET<="' + UsDateTime(DateFin) + '"';
        TobEvolutions := Tob.Create('LesEvolutions', nil, -1);
        TobEvolutions.LoadDetailDbFromSQL('LesEvolutions', St);
//        Ferme(Q);
//        if TobEvolutions.Detail.Count > 0 then
        For i := 0 to tobEvolutions.detail.count - 1 do
        begin
          NbMess := 0;
//          IDEtab := TobEvolutions.Detail[i].GetValue('PE2_ETABLISSEMENT');
          CodeActSal := TobEvolutions.Detail[i].GetValue('PE2_MSAACTIVITE');
          IdSal := Copy(TobEvolutions.Detail[i].GetValue('PSA_NUMEROSS'), 1, 13);
          MessError := ControleFichierMSA(IdSal, 215);
          if MessError <> '' then
          begin
            NbMess := NbMess + 1;
            TabMess[NbMess] := MessError;
          end;
          Nom := Copy(TobEvolutions.Detail[i].GetValue('PSA_LIBELLE'), 1, 25);
          Prenom := Copy(TobEvolutions.Detail[i].GetValue('PSA_PRENOM'), 1, 20);
          DateNaiss := TobEvolutions.Detail[i].GetValue('PSA_DATENAISSANCE');
          If DateNaiss > IDate1900 then SDateNaiss := FormatDateTime('yyyymmdd', DateNaiss)
          else SDateNaiss := '';
          MessError := ControleFichierMSA(SDateNaiss, 218);
          if MessError <> '' then
          begin
            NbMess := NbMess + 1;
            TabMess[NbMess] := MessError;
          end;
          DebAct := TobEvolutions.Detail[i].GetValue('PE2_DEBACTIVITE');
          if DebAct <> IDate1900 then
          begin
            SDebAct := FormatDateTime('yyyymmdd', DebAct);
          end
          else SDebAct := '';
          FinAct := TobEvolutions.Detail[i].GetValue('PE2_FINACTIVITE');
          if FinAct <> IDate1900 then
          begin
            SFinAct := FormatDateTime('yyyymmdd', FinAct);
          end
          else SFinAct := '';
          EcritureLigne(IdEntreprise, 'S', 13);
          EcritureLigne(IdEtab, 'S', 14);
          EcritureLigne(CodeActSal, 'I', 4);
          EcritureLigne('', 'S', 2);
          EcritureLigne('PE21', 'S', 4);
          EcritureLigne(IdSal, 'S', 13);
          EcritureLigne(Nom, 'S', 25);
          EcritureLigne(Prenom, 'S', 20);
          EcritureLigne(SDateNaiss, 'I', 8);
          EcritureLigne(SDebAct, 'I', 8);
          EcritureLigne(SFinAct, 'I', 8);
          EcritureLigne('', 'S', 81);
          WriteLN(FEcrt, '');
          if NbMess >= 1 then
          begin
            Writeln(FZ, '******************* Salarié ' + Nom + ' ' + Prenom +
              ', période du' + dateToStr(DateDebSal) +
              'au' + DateToStr(DateFinSal));
            Erreur := True;
          end;
          for j := 1 to NbMess do
          begin
            writeln(FZ, TabMess[j]);
          end;
          NbEnreg := NbEnreg + 1;
        end;
        FreeAndNil(TobEvolutions);
        //SEGMENT PE22
//Optimisation        Q := OpenSQL('SELECT MSAEVOLUTIONSPE2.*,PSA_LIBELLE,PSA_PRENOM,PSA_ETABLISSEMENT,PSA_NUMEROSS,PSA_DATENAISSANCE FROM MSAEVOLUTIONSPE2 ' +
//          'LEFT JOIN SALARIES ON PE2_SALARIE=PSA_SALARIE ' +
//          'WHERE PE2_TYPEEVOLMSA="PE22" ' +
//          'AND PE2_MSAACTIVITE="' + ActiviteMSA + '" AND PE2_ETABLISSEMENT="' + Etablissement + '" ' +
//          'AND PE2_DATEEFFET>="' + UsdateTime(DateDebut) + '" AND PE2_DATEEFFET<="' + UsDateTime(DateFin) + '"', True);
//        TobEvolutions.LoadDetailDB('LesEvolutions', '', '', Q, false);
        St := 'SELECT MSAEVOLUTIONSPE2.*,PSA_LIBELLE,PSA_PRENOM,PSA_ETABLISSEMENT,PSA_NUMEROSS,PSA_DATENAISSANCE FROM MSAEVOLUTIONSPE2 ' +
          'LEFT JOIN SALARIES ON PE2_SALARIE=PSA_SALARIE ' +
          'WHERE PE2_TYPEEVOLMSA="PE22" ' +
          'AND PE2_MSAACTIVITE="' + ActiviteMSA + '" AND PE2_ETABLISSEMENT="' + Etablissement + '" ' +
          'AND PE2_DATEEFFET>="' + UsdateTime(DateDebut) + '" AND PE2_DATEEFFET<="' + UsDateTime(DateFin) + '"';
        TobEvolutions := Tob.Create('Lesvolutions', nil, -1);
        TobEvolutions.LoadDetailDbFromSQL('LesEvolutions', St);
//        Ferme(Q);
//        if TobEvolutions.Detail.Count > 0 then
For i := 0 to tobEvolutions.detail.count - 1 do
        begin
          NbMess := 0;
//          IDEtab := TobEvolutions.Detail[i].GetValue('PE2_ETABLISSEMENT');
          CodeActSal := TobEvolutions.Detail[i].GetValue('PE2_MSAACTIVITE');
          IdSal := Copy(TobEvolutions.Detail[i].GetValue('PSA_NUMEROSS'), 1, 13);
          MessError := ControleFichierMSA(IdSal, 215);
          if MessError <> '' then
          begin
            NbMess := NbMess + 1;
            TabMess[NbMess] := MessError;
          end;
          Nom := Copy(TobEvolutions.Detail[i].GetValue('PSA_LIBELLE'), 1, 25);
          Prenom := Copy(TobEvolutions.Detail[i].GetValue('PSA_PRENOM'), 1, 20);
          DateNaiss := TobEvolutions.Detail[i].GetValue('PSA_DATENAISSANCE');
          If DateNaiss > IDate1900 then SDateNaiss := FormatDateTime('yyyymmdd', DateNaiss)
          else SDateNaiss := '';
          MessError := ControleFichierMSA(SDateNaiss, 218);
          if MessError <> '' then
          begin
            NbMess := NbMess + 1;
            TabMess[NbMess] := MessError;
          end;
          DebSuspCt := TobEvolutions.Detail[i].GetValue('PE2_DEBSUSPCT');
          if DebSuspCt <> IDate1900 then
          begin
            SDebSuspCt := FormatDateTime('yyyymmdd', DebSuspCt);
          end
          else SDebSuspCt := '';
          FinSuspCt := TobEvolutions.Detail[i].GetValue('PE2_FINSUSSPCT');
          if FinSuspCt <> IDate1900 then
          begin
            SFinSuspCt := FormatDateTime('yyyymmdd', FinSuspCt);
          end
          else SFinSuspCt := '';
          CodeSuspCt := TobEvolutions.Detail[i].GetValue('PE2_MSASUSPCT');
          EcritureLigne(IdEntreprise, 'S', 13);
          EcritureLigne(IdEtab, 'S', 14);
          EcritureLigne(CodeActSal, 'I', 4);
          EcritureLigne('', 'S', 2);
          EcritureLigne('PE22', 'S', 4);
          EcritureLigne(IdSal, 'S', 13);
          EcritureLigne(Nom, 'S', 25);
          EcritureLigne(Prenom, 'S', 20);
          EcritureLigne(SDateNaiss, 'I', 8);
          EcritureLigne(SDebSuspCt, 'I', 8);
          EcritureLigne(SFinSuspCt, 'I', 8);
          EcritureLigne(CodeSuspCt, 'I', 2);
          EcritureLigne('', 'S', 79);
          if NbMess >= 1 then
          begin
            Writeln(FZ, '******************* Salarié ' + Nom + ' ' + Prenom +
              ', période du' + dateToStr(DateDebSal) +
              'au' + DateToStr(DateFinSal));
            Erreur := True;
          end;
          for j := 1 to NbMess do
          begin
            writeln(FZ, TabMess[j]);
          end;
          WriteLN(FEcrt, '');
          NbEnreg := NbEnreg + 1;
        end;
        FreeAndNil(TobEvolutions);
        //SEGMENT PE23
//Optimisation        Q := OpenSQL('SELECT MSAEVOLUTIONSPE2.*,PSA_LIBELLE,PSA_PRENOM,PSA_ETABLISSEMENT,PSA_NUMEROSS,PSA_DATENAISSANCE FROM MSAEVOLUTIONSPE2 ' +
//          'LEFT JOIN SALARIES ON PE2_SALARIE=PSA_SALARIE ' +
//          'WHERE PE2_TYPEEVOLMSA="PE23" ' +
//          'AND PE2_MSAACTIVITE="' + ActiviteMSA + '" AND PE2_ETABLISSEMENT="' + Etablissement + '" ' +
//          'AND PE2_DATEEFFET>="' + UsdateTime(DateDebut) + '" AND PE2_DATEEFFET<="' + UsDateTime(DateFin) + '"', True);
//        TobEvolutions.LoadDetailDB('LesEvolutions', '', '', Q, false);
        St := 'SELECT MSAEVOLUTIONSPE2.*,PSA_LIBELLE,PSA_PRENOM,PSA_ETABLISSEMENT,PSA_NUMEROSS,PSA_DATENAISSANCE FROM MSAEVOLUTIONSPE2 ' +
          'LEFT JOIN SALARIES ON PE2_SALARIE=PSA_SALARIE ' +
          'WHERE PE2_TYPEEVOLMSA="PE23" ' +
          'AND PE2_MSAACTIVITE="' + ActiviteMSA + '" AND PE2_ETABLISSEMENT="' + Etablissement + '" ' +
          'AND PE2_DATEEFFET>="' + UsdateTime(DateDebut) + '" AND PE2_DATEEFFET<="' + UsDateTime(DateFin) + '"';
        TobEvolutions := Tob.Create('Lesvolutions', nil, -1);
        TobEvolutions.LoadDetailDbFromSQL('LesEvolutions', St);
//        Ferme(Q);
//        if TobEvolutions.Detail.Count > 0 then
For i := 0 to tobEvolutions.detail.count - 1 do
        begin
          NbMess := 0;
//          IDEtab := TobEvolutions.Detail[i].GetValue('PE2_ETABLISSEMENT');
          CodeActSal := TobEvolutions.Detail[i].GetValue('PE2_MSAACTIVITE');
          IdSal := Copy(TobEvolutions.Detail[i].GetValue('PSA_NUMEROSS'), 1, 13);
          MessError := ControleFichierMSA(IdSal, 215);
          if MessError <> '' then
          begin
            NbMess := NbMess + 1;
            TabMess[NbMess] := MessError;
          end;
          Nom := Copy(TobEvolutions.Detail[i].GetValue('PSA_LIBELLE'), 1, 25);
          Prenom := Copy(TobEvolutions.Detail[i].GetValue('PSA_PRENOM'), 1, 20);
          Prenom := PGUpperCase(Prenom);
          DateNaiss := TobEvolutions.Detail[i].GetValue('PSA_DATENAISSANCE');
          If DateNaiss > IDate1900 then SDateNaiss := FormatDateTime('yyyymmdd', DateNaiss)
          else SDateNaiss := '';
          MessError := ControleFichierMSA(SDateNaiss, 218);
          if MessError <> '' then
          begin
            NbMess := NbMess + 1;
            TabMess[NbMess] := MessError;
          end;
          DateEffet := TobEvolutions.Detail[i].GetValue('PE2_DATEEFFET');
          SDateEffet := FormatDateTime('yyyymmdd', DateEffet);
          DepartLieuTrav := Copy(TobEvolutions.Detail[i].GetValue('PE2_CPLIEUTRAV'), 1, 2);
          CommLieuTrav := Copy(TobEvolutions.Detail[i].GetValue('PE2_CPLIEUTRAV'), 3, 3);
          If TobEvolutions.Detail[i].GetValue('PE2_PCTTPSPART') > 0 then PTpsPart := TobEvolutions.Detail[i].GetValue('PE2_PCTTPSPART') * 100
          else PTpsPart := '';
          CodeTpsPart := TobEvolutions.Detail[i].GetValue('PE2_MSATPSPART');
          CodeCDD := TobEvolutions.Detail[i].GetValue('PE2_MSACONTRAT');
          If CodeCDD = 'CDI' then CodeCDD := 'I';
          If CodeCDD = 'CCD' then CodeCDD := 'D';
          NbHConvention := TobEvolutions.Detail[i].GetValue('PE2_NBHCONV');
          if NbHConvention <> '0' then NbHConvention := FloatToStr(StrToFloat(NbHConvention) * 100);
          Niveau := TobEvolutions.Detail[i].GetValue('PE2_MSANIVEAU');
          EcritureLigne(IdEntreprise, 'S', 13);
          EcritureLigne(IdEtab, 'S', 14);
          EcritureLigne(CodeActSal, 'I', 4);
          EcritureLigne('', 'S', 2);
          EcritureLigne('PE23', 'S', 4);
          EcritureLigne(IdSal, 'S', 13);
          EcritureLigne(Nom, 'S', 25);
          EcritureLigne(Prenom, 'S', 20);
          EcritureLigne(SDateNaiss, 'I', 8);
          EcritureLigne(SDateEffet, 'I', 8);
          EcritureLigne('', 'S', 3);
          EcritureLigne(DepartLieuTrav, 'I', 2);
          EcritureLigne(CommLieuTrav, 'I', 3);
          EcritureLigne(PTpsPart, 'I', 4);
          EcritureLigne(CodeTpsPart, 'I', 1);
          EcritureLigne(CodeCDD, 'S', 1);
          EcritureLigne(NbHConvention, 'I', 5);
          EcritureLigne(Niveau, 'S', 5);
          //DEB PT14
          PersTech := TobEvolutions.Detail[i].GetValue('PE2_PERSTECH');   //Périodicité de la durée de travail
          Qualite := TobEvolutions.Detail[i].GetValue('PE2_CODEQUALITE'); //Durée du contrat de travail
          EcritureLigne(PersTech, 'S', 1);
          EcritureLigne(Qualite, 'I', 6);
          //FIN PT14
          EcritureLigne('', 'S', 58);
          if NbMess >= 1 then
          begin
            Writeln(FZ, '******************* Salarié ' + Nom + ' ' + Prenom +
              ', période du' + dateToStr(DateDebSal) +
              'au' + DateToStr(DateFinSal));
            Erreur := True;
          end;
          for j := 1 to NbMess do
          begin
            writeln(FZ, TabMess[j]);
          end;
          WriteLN(FEcrt, '');
          NbEnreg := NbEnreg + 1;
        end;
        FreeAndNil(TobEvolutions);
        //SEGMENT PE24
//Optimisation        Q := OpenSQL('SELECT MSAEVOLUTIONSPE2.*,PSA_LIBELLE,PSA_PRENOM,PSA_ETABLISSEMENT,PSA_NUMEROSS,PSA_DATENAISSANCE FROM MSAEVOLUTIONSPE2 ' +
//          'LEFT JOIN SALARIES ON PE2_SALARIE=PSA_SALARIE ' +
//          'WHERE PE2_TYPEEVOLMSA="PE24" ' +
//          'AND PE2_MSAACTIVITE="' + ActiviteMSA + '" AND PE2_ETABLISSEMENT="' + Etablissement + '" ' +
//          'AND PE2_DATEEFFET>="' + UsdateTime(DateDebut) + '" AND PE2_DATEEFFET<="' + UsDateTime(DateFin) + '"', True);
//        TobEvolutions.LoadDetailDB('LesEvolutions', '', '', Q, false);
        St := 'SELECT MSAEVOLUTIONSPE2.*,PSA_LIBELLE,PSA_PRENOM,PSA_ETABLISSEMENT,PSA_NUMEROSS,PSA_DATENAISSANCE FROM MSAEVOLUTIONSPE2 ' +
          'LEFT JOIN SALARIES ON PE2_SALARIE=PSA_SALARIE ' +
          'WHERE PE2_TYPEEVOLMSA="PE24" ' +
          'AND PE2_MSAACTIVITE="' + ActiviteMSA + '" AND PE2_ETABLISSEMENT="' + Etablissement + '" ' +
          'AND PE2_DATEEFFET>="' + UsdateTime(DateDebut) + '" AND PE2_DATEEFFET<="' + UsDateTime(DateFin) + '"';
        TobEvolutions := Tob.Create('Lesvolutions', nil, -1);
        TobEvolutions.LoadDetailDbFromSQL('LesEvolutions', St);
//        Ferme(Q);
//        if TobEvolutions.Detail.Count > 0 then
        For i := 0 to tobEvolutions.detail.count - 1 do
        begin
                NbMess := 0;
//          IDEtab := TobEvolutions.Detail[i].GetValue('PE2_ETABLISSEMENT');
                CodeActSal := TobEvolutions.Detail[i].GetValue('PE2_MSAACTIVITE');
                IdSal := Copy(TobEvolutions.Detail[i].GetValue('PSA_NUMEROSS'), 1, 13);
                MessError := ControleFichierMSA(IdSal, 215);
                if MessError <> '' then
                begin
                        NbMess := NbMess + 1;
                        TabMess[NbMess] := MessError;
                end;
                Nom := Copy(TobEvolutions.Detail[i].GetValue('PSA_LIBELLE'), 1, 25);
                Prenom := Copy(TobEvolutions.Detail[i].GetValue('PSA_PRENOM'), 1, 20);
                Prenom := PGUpperCase(Prenom);
                DateNaiss := TobEvolutions.Detail[i].GetValue('PSA_DATENAISSANCE');
                If DateNaiss > IDate1900 then SDateNaiss := FormatDateTime('yyyymmdd', DateNaiss)
                else SDateNaiss := '';
                MessError := ControleFichierMSA(SDateNaiss, 218);
                if MessError <> '' then
                begin
                        NbMess := NbMess + 1;
                        TabMess[NbMess] := MessError;
                end;
                DateEffet := TobEvolutions.Detail[i].GetValue('PE2_DATEEFFET');
                SDateEffet := FormatDateTime('yyyymmdd', DateEffet);
                PersTech := TobEvolutions.Detail[i].GetValue('PE2_PERSTECH');
                if PersTech = 'OUI' then PersTech := 'O';
                if PersTech = 'NON' then PersTech := 'N';
                TopCadre := TobEvolutions.Detail[i].GetValue('PE2_TOPCADRE');
                if TopCadre = 'OUI' then TopCadre := 'O';
                if TopCadre = 'NON' then TopCadre := 'N';
                PolyEmp := TobEvolutions.Detail[i].GetValue('PE2_MSAPOLYEMP');
                PersTechCUMA := TobEvolutions.Detail[i].GetValue('PE2_PERSTECHCUMA');
                if PersTechCUMA = 'OUI' then PersTechCUMA := 'O';
                if PersTechCUMA = 'NON' then PersTechCUMA := 'N';
                FiscHorsF := TobEvolutions.Detail[i].GetValue('PE2_FISCHORSF');
                if FiscHorsF = 'OUI' then FiscHorsF := 'O';
                if FiscHorsF = 'NON' then FiscHorsF := 'N';
                CodeClasse := TobEvolutions.Detail[i].GetValue('PE2_CLASSEELEVE');
                UG := TobEvolutions.Detail[i].GetValue('PE2_MSAUG');
                Qualite := TobEvolutions.Detail[i].GetValue('PE2_CODEQUALITE');
                CRCCA := TobEvolutions.Detail[i].GetValue('PE2_MSACODECRCCA');
                If CRCCA = '01' then CRCCA := '03';
                If CRCCA = '02' then CRCCA := '36';
                AUBRY1 := TobEvolutions.Detail[i].GetValue('PE2_MSAAUBRY1');
                if AUBRY1 = 'OUI' then AUBRY1 := 'O';
                if AUBRY1 = 'NON' then AUBRY1 := 'N';
                AUBRY2 := TobEvolutions.Detail[i].GetValue('PE2_MSAAUBRY2');
                if AUBRY2 = 'OUI' then AUBRY2 := 'O';
                if AUBRY2 = 'NON' then AUBRY2 := 'N';
                Emploi := TobEvolutions.Detail[i].GetValue('PE2_LIBELLEEMPLOI');
                Retraite := TobEvolutions.Detail[i].GetValue('PE2_RETRAITEACT');
                if Retraite = 'OUI' then Retraite := 'O';
                if Retraite = 'NON' then Retraite := 'N';
                Convention := TobEvolutions.Detail[i].GetValue('PE2_MSACONVCOLL');
                SectionP := TobEvolutions.Detail[i].GetValue('PE2_MSATPSPART');
//PT14                if (SectionP <> '') then SectionP := '0'+SectionP;
                CollegeP  := TobEvolutions.Detail[i].GetValue('PE2_MSACONTRAT');
//PT14                if (CollegeP <> '') then CollegeP := '0'+CollegeP;
                //          if convention <> '' then Convention := RechDom('PGCONVENTION', Convention, False);
                EcritureLigne(IdEntreprise, 'S', 13);
                EcritureLigne(IdEtab, 'S', 14);
                EcritureLigne(CodeActSal, 'I', 4);
                EcritureLigne('', 'S', 2);
                EcritureLigne('PE24', 'S', 4);
                EcritureLigne(IdSal, 'S', 13);
                EcritureLigne(Nom, 'S', 25);
                EcritureLigne(Prenom, 'S', 20);
                EcritureLigne(SDateNaiss, 'I', 8);
                EcritureLigne(SDateEffet, 'I', 8);
                EcritureLigne(PersTech, 'S', 1);
                EcritureLigne('', 'S', 1);
                EcritureLigne(TopCadre, 'S', 1);
                EcritureLigne('', 'S', 1);
                EcritureLigne(PolyEmp, 'I', 1);
                EcritureLigne(PersTechCUMA, 'S', 1);
                EcritureLigne(FiscHorsF, 'S', 1);
                EcritureLigne(CodeClasse, 'I', 1);
                EcritureLigne(UG, 'S', 2);
                EcritureLigne(Qualite, 'I', 3);
                EcritureLigne(CRCCA, 'I', 2);
                EcritureLigne(AUBRY1, 'S', 1);
                EcritureLigne(AUBRY2, 'S', 1);
                EcritureLigne(Emploi, 'S', 25);
                EcritureLigne(Retraite, 'S', 1);
                EcritureLigne(Convention, 'S', 25);
                EcritureLigne(SectionP, 'I', 2);    //PT14
                EcritureLigne(CollegeP, 'I', 2);    //PT14
                EcritureLigne('', 'S', 17);
                WriteLN(FEcrt, '');
                NbEnreg := NbEnreg + 1;
        end;
        FreeAndNil(TobEvolutions);
        For u := 0 to TobUG.Detail.Count -1 do
        begin
               UniteGestion := TobUG.Detail[u].GetValue('UG');
               //SEGMENT PE31
               Q := OpenSQL('SELECT COUNT (DISTINCT PE3_SALARIE) NBSAL FROM MSAPERIODESPE31 WHERE' +
               ' PE3_DATEDEBUT>="' + UsDatetime(Datedebut) + '"' +
               ' AND PE3_DATEFIN<="' + UsDatetime(DateFin) + '"' +
               ' AND PE3_MSAACTIVITE="' + ActiviteMSA + '"' +
               ' AND PE3_UGMSA="' + UniteGestion + '"' +
               ' AND PE3_ETABLISSEMENT="' + Etablissement + '"' +
               ' AND PE3_ORDRE < 100', True);
               if not Q.Eof then NbSalaries := NbSalaries + Q.FindField('NBSAL').AsInteger;
               Ferme(Q);
               St := 'SELECT * FROM MSAPERIODESPE31 WHERE' +
               ' PE3_DATEDEBUT>="' + UsDatetime(Datedebut) + '"' +
               ' AND PE3_DATEFIN<="' + UsDatetime(DateFin) + '"' +
               ' AND PE3_MSAACTIVITE="' + ActiviteMSA + '"' +
               ' AND PE3_UGMSA="' + UniteGestion + '"' +
               ' AND PE3_ETABLISSEMENT="' + Etablissement + '"' +
               ' AND PE3_ORDRE < 100';
               TobMSA := Tob.create('MSA', nil, -1);
               TobMSA.LoadDetailDbFromSQL('MSA', St);
               DeclarerAsiette := False;
               //DEB PT18
               MemSalarie := '';
               MemDateDebut := iDate1900;
               MemDateFin := iDate1900;
               //FIN PT18
               for i := 0 to TobMSA.Detail.Count - 1 do
               begin
                    //DEB PT18
                    MemSalarie := TobMSA.Detail[i].GetValue('PE3_SALARIE');
                    MemDateDebut := TobMSA.Detail[i].GetValue('PE3_DATEDEBUT');
                    MemDateFin := TobMSA.Detail[i].GetValue('PE3_DATEFIN');
                    //FIN PT18
                    DeclarerAsiette := True;
                    NbMess := 0;
                    CodeActSal := TobMSA.Detail[i].GetValue('PE3_MSAACTIVITE');
                    UG := TobMSA.Detail[i].GetValue('PE3_UGMSA');
                    IdSal := copy(TobMSA.Detail[i].GetValue('PE3_NUMEROSS'), 1, 13);
                    Nom := copy(TobMSA.Detail[i].GetValue('PE3_NOM'), 1, 25);
                    Prenom := copy(TobMSA.Detail[i].GetValue('PE3_PRENOM'), 1, 20);
                    if TobMSA.Detail[i].GetValue('PE3_TOPPLAF') = 'X' then TopNdP := 'O'
                    else TopNdP := 'N';
                    PourcExo := TobMSA.Detail[i].GetValue('PE3_PEXOMSA');
                    NbJ := TobMSA.Detail[i].GetValue('PE3_NBJOURS');
                    //          NbH := FloatToStr(TobMSA.Detail[i].GetValue('PE3_NBHEURES'));
                    NbH := StrfMontant(TobMSA.Detail[i].GetValue('PE3_NBHEURES'), 15, 2, '', TRUE);
                    EltC1 := TobMSA.Detail[i].GetValue('PE3_ELTCALCMSA1');
                    MessError := ControleFichierMSA(IdSal, 215);
                    if MessError <> '' then
                    begin
                         NbMess := NbMess + 1;
                         TabMess[NbMess] := MessError;
                    end;
                    MessError := ControleFichierMSA(EltC1, 3115);
                    if MessError <> '' then
                    begin
                         NbMess := NbMess + 1;
                         TabMess[NbMess] := MessError;
                    end;
                    EltC2 := TobMSA.Detail[i].GetValue('PE3_ELTCALCMSA2');
                    EltC3 := TobMSA.Detail[i].GetValue('PE3_ELTCALCMSA3');
                    EltC4 := TobMSA.Detail[i].GetValue('PE3_ELTCALCMSA4');
                    EltC5 := TobMSA.Detail[i].GetValue('PE3_ELTCALCMSA5');
                    Velt1 := FloatToStr(Arrondi(TobMSA.Detail[i].GetValue('PE3_MONTANTELT1'), 0));
                    Velt2 := FloatToStr(Arrondi(TobMSA.Detail[i].GetValue('PE3_MONTANTELT2'), 0));
                    Velt3 := FloatToStr(Arrondi(TobMSA.Detail[i].GetValue('PE3_MONTANTELT3'), 0));
                    Velt4 := FloatToStr(Arrondi(TobMSA.Detail[i].GetValue('PE3_MONTANTELT4'), 0));
                    Velt5 := FloatToStr(Arrondi(TobMSA.Detail[i].GetValue('PE3_MONTANTELT5'), 0));
                    DateNaiss := TobMSA.Detail[i].GetValue('PE3_DATENAISSANCE');
                    DateDebSal := TobMSA.Detail[i].GetValue('PE3_DATEDEBUT');
                    DateFinSal := TobMSA.Detail[i].GetValue('PE3_DATEFIN');
                    If DateNaiss > IDate1900 then SDateNaiss := FormatDateTime('yyyymmdd', DateNaiss)
                    else SDateNaiss := '';
                    MessError := ControleFichierMSA(SDateNaiss, 218);
                    if MessError <> '' then
                    begin
                         NbMess := NbMess + 1;
                         TabMess[NbMess] := MessError;
                    end;
                    SDateDebSal := FormatDateTime('yyyymmdd', DateDebSal);
                    SDateFinSal := FormatDateTime('yyyymmdd', DateFinSal);
                    Prenom := PGUpperCase(Prenom);
                    Prenom := ControlePonctuation(Prenom);
                    MessError := ControleFichierMSA(Prenom, 7);
                    if MessError <> '' then
                    begin
                         NbMess := NbMess + 1;
                         TabMess[NbMess] := MessError;
                    end;
                    EcritureLigne(IdEntreprise, 'S', 13);
                    EcritureLigne(IdEtab, 'S', 14);
                    EcritureLigne(CodeActSal, 'I', 4);
                    EcritureLigne(UG, 'S', 2);
                    EcritureLigne('PE31', 'S', 4);
                    EcritureLigne(IdSal, 'I', 13);
                    EcritureLigne(Nom, 'S', 25);
                    EcritureLigne(Prenom, 'S', 20);
                    EcritureLigne(SDateNaiss, 'I', 8);
                    EcritureLigne(SDateDebSal, 'I', 8);
                    EcritureLigne(SDateFinSal, 'I', 8);
                    EcritureLigne(TopNdP, 'S', 1);
                    EcritureLigne(PourcExo, 'I', 1);
                    EcritureLigne('   ', 'S', 3);
                    EcritureLigne(NbJ, 'I', 3);
                    EcritureLigne(NbH, 'H', 5);
                    EcritureLigne(Eltc1, 'I', 2);
                    EcritureLigne(Velt1, 'I', 6);
                    EcritureLigne(EltC2, 'I', 2);
                    EcritureLigne(Velt2, 'I', 6);
                    EcritureLigne(EltC3, 'I', 2);
                    EcritureLigne(Velt3, 'I', 6);
                    EcritureLigne(EltC4, 'I', 2);
                    EcritureLigne(Velt4, 'I', 6);
                    EcritureLigne(EltC5, 'I', 2);
                    EcritureLigne(Velt5, 'I', 6);
                    EcritureLigne('', 'S', 28);
                    if (EltC1 = '01') or (EltC1 = '02') or (EltC1 = '09') or (EltC1 = '12') or (EltC1 = '14') then
                      MontantSalairesEtab := MontantSalairesEtab + StrToFloat(Velt1);
                    if (EltC2 = '01') or (EltC2 = '02') or (EltC2 = '09') or (EltC2 = '12') or (EltC2 = '14') then
                      MontantSalairesEtab := MontantSalairesEtab + StrToFloat(Velt2);
                    if (EltC3 = '01') or (EltC3 = '02') or (EltC3 = '09') or (EltC3 = '12') or (EltC3 = '14') then
                      MontantSalairesEtab := MontantSalairesEtab + StrToFloat(Velt3);
                    if (EltC4 = '01') or (EltC4 = '02') or (EltC4 = '09') or (EltC4 = '12') or (EltC4 = '14') then
                      MontantSalairesEtab := MontantSalairesEtab + StrToFloat(Velt4);
                    if (EltC5 = '01') or (EltC5 = '02') or (EltC5 = '09') or (EltC5 = '12') or (EltC5 = '14') then
                      MontantSalairesEtab := MontantSalairesEtab + StrToFloat(Velt5);
                    if NbMess >= 1 then
                    begin
                      Writeln(FZ, '******************* Salarié ' + Nom + ' ' + Prenom +
                        ', période du' + dateToStr(DateDebSal) +
                        'au' + DateToStr(DateFinSal));
                      Erreur := True;
                    end;
                    for j := 1 to NbMess do
                    begin
                         writeln(FZ, TabMess[j]);
                    end;
                    WriteLN(FEcrt, '');
                    NbEnreg := NbEnreg + 1;
                    if (EltC1 = '01') or (EltC1 = '02') or (EltC1 = '09') or (EltC1 = '12') or (EltC1 = '14') then
                      MontantSalairesUG := MontantSalairesUG + StrToFloat(Velt1);
                    if (EltC2 = '01') or (EltC2 = '02') or (EltC2 = '09') or (EltC2 = '12') or (EltC2 = '14') then
                      MontantSalairesUG := MontantSalairesUG + StrToFloat(Velt2);
                    if (EltC3 = '01') or (EltC3 = '02') or (EltC3 = '09') or (EltC3 = '12') or (EltC3 = '14') then
                      MontantSalairesUG := MontantSalairesUG + StrToFloat(Velt3);
                    if (EltC4 = '01') or (EltC4 = '02') or (EltC4 = '09') or (EltC4 = '12') or (EltC4 = '14') then
                      MontantSalairesUG := MontantSalairesUG + StrToFloat(Velt4);
                    if (EltC5 = '01') or (EltC5 = '02') or (EltC5 = '09') or (EltC5 = '12') or (EltC5 = '14') then
                      MontantSalairesUG := MontantSalairesUG + StrToFloat(Velt5);

                   //DEB PT18
                   if ((i<TobMSA.Detail.Count-1)
                      and ((MemSalarie <> TobMSA.Detail[i+1].GetValue('PE3_SALARIE'))
                      or (MemDateDebut <> TobMSA.Detail[i+1].GetValue('PE3_DATEDEBUT'))
                      or (MemDateFin <> TobMSA.Detail[i+1].GetValue('PE3_DATEFIN'))))
                      or (i=TobMSA.Detail.Count-1) then
                   begin
                     //Récupérer le code PE32 au fur et à mesure pour chaque salarié et par période
                     St := 'SELECT * FROM MSAPERIODESPE31 WHERE' +
                     ' PE3_SALARIE="' + TobMSA.Detail[i].GetValue('PE3_SALARIE') + '" ' +
                     ' AND PE3_DATEDEBUT="' + UsDatetime(DateDebSal) + '"' +
                     ' AND PE3_DATEFIN<="' + UsDatetime(DateFinSal) + '"' +
                     ' AND PE3_MSAACTIVITE="' + ActiviteMSA + '"' +
                     ' AND PE3_UGMSA="' + UniteGestion + '"' +
                     ' AND PE3_ETABLISSEMENT="' + Etablissement + '"' +
                     ' AND PE3_ORDRE >= 100' +
                     ' ORDER BY PE3_ORDRE';
                     TobMSA2 := Tob.create('MSA', nil, -1);
                     TobMSA2.LoadDetailDbFromSQL('MSA', St);
                     for r := 1 to 4 do
                     begin
                       TypeRem[r] := '';
                       TauxMajo[r] := '';
                       NbHeures[r] := '';
                       MontantRem[r] := '';
                     end;
                     if TobMSA2.Detail.Count > 0 then
                     begin
                       z := 0;
                       while z < TobMSA2.Detail.Count do
                       begin
                         TypeRem[z+1] := FloatToStr(Arrondi(TobMSA2.Detail[z].GetValue('PE3_MONTANTELT1'), 0));
                         TauxMajo[z+1] := FloatToStr(Arrondi(TobMSA2.Detail[z].GetValue('PE3_MONTANTELT2'), 0));
                         NbHeures[z+1] := FloatToStr(Arrondi(TobMSA2.Detail[z].GetValue('PE3_MONTANTELT3'), 0));
                         MontantRem[z+1] := FloatToStr(Arrondi(TobMSA2.Detail[z].GetValue('PE3_MONTANTELT4'), 0));

                         NbMess := 0;
                         CodeActSal := TobMSA2.Detail[z].GetValue('PE3_MSAACTIVITE');
                         UG := TobMSA2.Detail[z].GetValue('PE3_UGMSA');
                         IdSal := copy(TobMSA2.Detail[z].GetValue('PE3_NUMEROSS'), 1, 13);
                         Nom := copy(TobMSA2.Detail[z].GetValue('PE3_NOM'), 1, 25);
                         Prenom := copy(TobMSA2.Detail[z].GetValue('PE3_PRENOM'), 1, 20);
                         MessError := ControleFichierMSA(IdSal, 215);
                         if MessError <> '' then
                         begin
                           NbMess := NbMess + 1;
                           TabMess[NbMess] := MessError;
                         end;
                         DateNaiss := TobMSA2.Detail[z].GetValue('PE3_DATENAISSANCE');
                         DateDebSal := TobMSA2.Detail[z].GetValue('PE3_DATEDEBUT');
                         DateFinSal := TobMSA2.Detail[z].GetValue('PE3_DATEFIN');
                         If DateNaiss > IDate1900 then SDateNaiss := FormatDateTime('yyyymmdd', DateNaiss)
                         else SDateNaiss := '';
                         MessError := ControleFichierMSA(SDateNaiss, 218);
                         if MessError <> '' then
                         begin
                           NbMess := NbMess + 1;
                           TabMess[NbMess] := MessError;
                         end;
                         SDateDebSal := FormatDateTime('yyyymmdd', DateDebSal);
                         SDateFinSal := FormatDateTime('yyyymmdd', DateFinSal);
                         Prenom := PGUpperCase(Prenom);
                         Prenom := ControlePonctuation(Prenom);
                         MessError := ControleFichierMSA(Prenom, 7);
                         if MessError <> '' then
                         begin
                           NbMess := NbMess + 1;
                           TabMess[NbMess] := MessError;
                         end;
                         EcritureLigne(IdEntreprise, 'S', 13);
                         EcritureLigne(IdEtab, 'S', 14);
                         EcritureLigne(CodeActSal, 'I', 4);
                         EcritureLigne(UG, 'S', 2);
                         EcritureLigne('PE32', 'S', 4);
                         EcritureLigne(IdSal, 'I', 13);
                         EcritureLigne(Nom, 'S', 25);
                         EcritureLigne(Prenom, 'S', 20);
                         EcritureLigne(SDateNaiss, 'I', 8);
                         EcritureLigne(SDateDebSal, 'I', 8);
                         EcritureLigne(SDateFinSal, 'I', 8);
                         EcritureLigne(TypeRem[1], 'I', 2);
                         EcritureLigne(TauxMajo[1], 'I', 3);
                         EcritureLigne(NbHeures[1], 'I', 5);
                         EcritureLigne(MontantRem[1], 'I', 6);
                         EcritureLigne(TypeRem[2], 'I', 2);
                         EcritureLigne(TauxMajo[2], 'I', 3);
                         EcritureLigne(NbHeures[2], 'I', 5);
                         EcritureLigne(MontantRem[2], 'I', 6);
                         EcritureLigne(TypeRem[3], 'I', 2);
                         EcritureLigne(TauxMajo[3], 'I', 3);
                         EcritureLigne(NbHeures[3], 'I', 5);
                         EcritureLigne(MontantRem[3], 'I', 6);
                         EcritureLigne(TypeRem[4], 'I', 2);
                         EcritureLigne(TauxMajo[4], 'I', 3);
                         EcritureLigne(NbHeures[4], 'I', 5);
                         EcritureLigne(MontantRem[4], 'I', 6);
                         EcritureLigne('', 'S', 17);

                         if NbMess >= 1 then
                         begin
                           Writeln(FZ, '******************* Salarié ' + Nom + ' ' + Prenom +
                             ', période du' + dateToStr(DateDebSal) +
                             'au' + DateToStr(DateFinSal));
                           Erreur := True;
                         end;
                         for j := 1 to NbMess do
                         begin
                           writeln(FZ, TabMess[j]);
                         end;
                         WriteLN(FEcrt, '');
                         z := z + 1;
                       end;
                     end;
                     FreeAndNil(TobMSA2);
                   end;
                   //FIN PT18
               end;
               FreeAndNil(TobMSA);

               //DEB PT14
               //SEGMENT PE32
{PT18               Q := OpenSQL('SELECT COUNT (DISTINCT PE3_SALARIE) NBSAL FROM MSAPERIODESPE31 WHERE' +
               ' PE3_DATEDEBUT>="' + UsDatetime(Datedebut) + '"' +
               ' AND PE3_DATEFIN<="' + UsDatetime(DateFin) + '"' +
               ' AND PE3_MSAACTIVITE="' + ActiviteMSA + '"' +
               ' AND PE3_UGMSA="' + UniteGestion + '"' +
               ' AND PE3_ETABLISSEMENT="' + Etablissement + '"' +
               ' AND PE3_ORDRE >= 100', True);
               if not Q.Eof then
               begin
//Optimisation                 Q := OpenSQL('SELECT * FROM MSAPERIODESPE31 WHERE' +
//                 ' PE3_DATEDEBUT>="' + UsDatetime(Datedebut) + '"' +
//                 ' AND PE3_DATEFIN<="' + UsDatetime(DateFin) + '"' +
//                 ' AND PE3_MSAACTIVITE="' + ActiviteMSA + '"' +
//                 ' AND PE3_UGMSA="' + UniteGestion + '"' +
//                 ' AND PE3_ETABLISSEMENT="' + Etablissement + '"' +
//                 ' AND PE3_ORDRE >= 100' +
//                 ' ORDER BY PE3_SALARIE,PE3_DATEDEBUT,PE3_DATEFIN,PE3_ORDRE', True);
//                 TobMSA.LoadDetailDB('MSA', '', '', Q, False);
                   St := 'SELECT * FROM MSAPERIODESPE31 WHERE' +
                   ' PE3_DATEDEBUT>="' + UsDatetime(Datedebut) + '"' +
                   ' AND PE3_DATEFIN<="' + UsDatetime(DateFin) + '"' +
                   ' AND PE3_MSAACTIVITE="' + ActiviteMSA + '"' +
                   ' AND PE3_UGMSA="' + UniteGestion + '"' +
                   ' AND PE3_ETABLISSEMENT="' + Etablissement + '"' +
                   ' AND PE3_ORDRE >= 100' +
                   ' ORDER BY PE3_SALARIE,PE3_DATEDEBUT,PE3_DATEFIN,PE3_ORDRE';
                   TobMSA := Tob.create('MSA', nil, -1);
                   TobMSA.LoadDetailDbFromSQL('MSA', St);
  //                 Ferme(Q);
                   MemSalarie := '';
                   MemDateDebut := iDate1900;
                   MemDateFin := iDate1900;
                   i := 0;
                   while i < TobMSA.Detail.Count do
                   begin
                        MemSalarie := TobMSA.Detail[i].GetValue('PE3_SALARIE');
                        MemDateDebut := TobMSA.Detail[i].GetValue('PE3_DATEDEBUT');
                        MemDateFin := TobMSA.Detail[i].GetValue('PE3_DATEFIN');
                        for r := 1 to 4 do
                        begin
                          TypeRem[r] := '';
                          TauxMajo[r] := '';
                          NbHeures[r] := '';
                          MontantRem[r] := '';
                        end;
                        r := 1;
                        while (MemSalarie = TobMSA.Detail[i].GetValue('PE3_SALARIE'))
                          and (MemDateDebut = TobMSA.Detail[i].GetValue('PE3_DATEDEBUT'))
                          and (MemDateFin = TobMSA.Detail[i].GetValue('PE3_DATEFIN')) do
                        begin
                          TypeRem[r] := FloatToStr(Arrondi(TobMSA.Detail[i].GetValue('PE3_MONTANTELT1'), 0));
                          TauxMajo[r] := FloatToStr(Arrondi(TobMSA.Detail[i].GetValue('PE3_MONTANTELT2'), 0));
                          NbHeures[r] := FloatToStr(Arrondi(TobMSA.Detail[i].GetValue('PE3_MONTANTELT3'), 0));
                          MontantRem[r] := FloatToStr(Arrondi(TobMSA.Detail[i].GetValue('PE3_MONTANTELT4'), 0));
                          r := r + 1;
                          i := i + 1;
                          if i >= TobMSA.Detail.Count then break;
                        end;
                        i := i - 1;
                        NbMess := 0;
                        CodeActSal := TobMSA.Detail[i].GetValue('PE3_MSAACTIVITE');
                        UG := TobMSA.Detail[i].GetValue('PE3_UGMSA');
                        IdSal := copy(TobMSA.Detail[i].GetValue('PE3_NUMEROSS'), 1, 13);
                        Nom := copy(TobMSA.Detail[i].GetValue('PE3_NOM'), 1, 25);
                        Prenom := copy(TobMSA.Detail[i].GetValue('PE3_PRENOM'), 1, 20);
                        MessError := ControleFichierMSA(IdSal, 215);
                        if MessError <> '' then
                        begin
                             NbMess := NbMess + 1;
                             TabMess[NbMess] := MessError;
                        end;
                        DateNaiss := TobMSA.Detail[i].GetValue('PE3_DATENAISSANCE');
                        DateDebSal := TobMSA.Detail[i].GetValue('PE3_DATEDEBUT');
                        DateFinSal := TobMSA.Detail[i].GetValue('PE3_DATEFIN');
                        If DateNaiss > IDate1900 then SDateNaiss := FormatDateTime('yyyymmdd', DateNaiss)
                        else SDateNaiss := '';
                        MessError := ControleFichierMSA(SDateNaiss, 218);
                        if MessError <> '' then
                        begin
                             NbMess := NbMess + 1;
                             TabMess[NbMess] := MessError;
                        end;
                        SDateDebSal := FormatDateTime('yyyymmdd', DateDebSal);
                        SDateFinSal := FormatDateTime('yyyymmdd', DateFinSal);
                        Prenom := PGUpperCase(Prenom);
                        Prenom := ControlePonctuation(Prenom);
                        MessError := ControleFichierMSA(Prenom, 7);
                        if MessError <> '' then
                        begin
                             NbMess := NbMess + 1;
                             TabMess[NbMess] := MessError;
                        end;
                        EcritureLigne(IdEntreprise, 'S', 13);
                        EcritureLigne(IdEtab, 'S', 14);
                        EcritureLigne(CodeActSal, 'I', 4);
                        EcritureLigne(UG, 'S', 2);
                        EcritureLigne('PE32', 'S', 4);
                        EcritureLigne(IdSal, 'I', 13);
                        EcritureLigne(Nom, 'S', 25);
                        EcritureLigne(Prenom, 'S', 20);
                        EcritureLigne(SDateNaiss, 'I', 8);
                        EcritureLigne(SDateDebSal, 'I', 8);
                        EcritureLigne(SDateFinSal, 'I', 8);
                        EcritureLigne(TypeRem[1], 'I', 2);
                        EcritureLigne(TauxMajo[1], 'I', 3);
                        EcritureLigne(NbHeures[1], 'I', 5);
                        EcritureLigne(MontantRem[1], 'I', 6);
                        EcritureLigne(TypeRem[2], 'I', 2);
                        EcritureLigne(TauxMajo[2], 'I', 3);
                        EcritureLigne(NbHeures[2], 'I', 5);
                        EcritureLigne(MontantRem[2], 'I', 6);
                        EcritureLigne(TypeRem[3], 'I', 2);
                        EcritureLigne(TauxMajo[3], 'I', 3);
                        EcritureLigne(NbHeures[3], 'I', 5);
                        EcritureLigne(MontantRem[3], 'I', 6);
                        EcritureLigne(TypeRem[4], 'I', 2);
                        EcritureLigne(TauxMajo[4], 'I', 3);
                        EcritureLigne(NbHeures[4], 'I', 5);
                        EcritureLigne(MontantRem[4], 'I', 6);
                        EcritureLigne('', 'S', 17);

                        if NbMess >= 1 then
                        begin
                          Writeln(FZ, '******************* Salarié ' + Nom + ' ' + Prenom +
                            ', période du' + dateToStr(DateDebSal) +
                            'au' + DateToStr(DateFinSal));
                          Erreur := True;
                        end;
                        for j := 1 to NbMess do
                        begin
                             writeln(FZ, TabMess[j]);
                        end;
                        WriteLN(FEcrt, '');
                        i := i + 1;
                   end;
                   FreeAndNil(TobMSA);
               end;
               Ferme(Q);
               //FIN PT14 }

               If TobUG.Detail.Count > 1 then
               begin
                    EltC1 := '';
                    VElt1 := '';
                    EltC2 := '';
                    VElt2 := '';
                    EltC3 := '';
                    VElt3 := '';
                    EltC4 := '';
                    VElt4 := '';
                    EltC5 := '';
                    VElt5 := '';
                    //Enregistrement PE41
                    if DeclarerAsiette = True then
                    begin
                        TotalSalaires := FloatToStr(Arrondi(MontantSalairesUG, 0));
                        TLA := TobLesAssiettes.FindFirst(['ETABLISSEMENT', 'ACTIVITE','UG'], [Etablissement, ActiviteMSA,UniteGestion], False);
                        if TLA <> nil then
                        begin
                          NumElement := 0;
                          For elt := 1 to 10 do
                          begin
                               TabElt[elt] := '';
                          end;
                          if TLA.GetValue('51') > 0 then
                          begin
                            NumElement := NumElement + 1;
                            TabElt[NumElement] := '51';
                          end;
                          if TLA.GetValue('53') > 0 then
                          begin
                            NumElement := NumElement + 1;
                            TabElt[NumElement] := '53';
                          end;
                          if TLA.GetValue('54') > 0 then
                          begin
                            NumElement := NumElement + 1;
                            TabElt[NumElement] := '54';
                          end;
                          if TLA.GetValue('56') > 0 then
                          begin
                            NumElement := NumElement + 1;
                            TabElt[NumElement] := '56';
                          end;
                          if TLA.GetValue('63') > 0 then
                          begin
                            NumElement := NumElement + 1;
                            TabElt[NumElement] := '63';
                          end;
                          //DEB PT18
                          if TLA.GetValue('66') > 0 then
                          begin
                            NumElement := NumElement + 1;
                            TabElt[NumElement] := '66';
                          end;
                          if TLA.GetValue('67') > 0 then
                          begin
                            NumElement := NumElement + 1;
                            TabElt[NumElement] := '67';
                          end;
                          if TLA.GetValue('68') > 0 then
                          begin
                            NumElement := NumElement + 1;
                            TabElt[NumElement] := '68';
                          end;
                          //FIN PT18
                          if TLA.GetValue('58') > 0 then
                          begin
                            NumElement := NumElement + 1;
                            TabElt[NumElement] := '58';
                          end;
                          if TLA.GetValue('59') > 0 then
                          begin
                            NumElement := NumElement + 1;
                            TabElt[NumElement] := '59';
                          end;
{PT18                          if TLA.GetValue('60') > 0 then
                          begin
                            NumElement := NumElement + 1;
                            TabElt[NumElement] := '60';
                          end;}
                          if TLA.GetValue('61') > 0 then
                          begin
                            NumElement := NumElement + 1;
                            TabElt[NumElement] := '61';
                          end;
                          if TLA.GetValue('62') > 0 then
                          begin
                            NumElement := NumElement + 1;
                            TabElt[NumElement] := '62';
                          end;
                          EltC1 := TabElt[1];
                          If EltC1 <> '' then VElt1 := FloatToStr(arrondi(TLA.GetValue(TabElt[1]), 0))
                          else VElt1 := '';
                          EltC2 := TabElt[2];
                          If EltC2 <> '' then VElt2 := FloatToStr(arrondi(TLA.GetValue(TabElt[2]), 0))
                          else VElt2 := '';
                          EltC3 := TabElt[3];
                          If EltC3 <> '' then VElt3 := FloatToStr(arrondi(TLA.GetValue(TabElt[3]), 0))
                          else VElt3 := '';
                          EltC4 := TabElt[4];
                          If EltC4 <> '' then VElt4 := FloatToStr(arrondi(TLA.GetValue(TabElt[4]), 0))
                          else VElt4 := '';
                          EltC5 := TabElt[5];
                          If EltC5 <> '' then VElt5 := FloatToStr(arrondi(TLA.GetValue(TabElt[5]), 0))
                          else VElt5 := '';
                          EltC6 := TabElt[6];
                          If EltC6 <> '' then VElt6 := FloatToStr(arrondi(TLA.GetValue(TabElt[6]), 0))
                          else VElt6 := '';
                          EltC7 := TabElt[7];
                          If EltC7 <> '' then VElt7 := FloatToStr(arrondi(TLA.GetValue(TabElt[7]), 0))
                          else VElt7 := '';
                          EltC8 := TabElt[8];
                          If EltC8 <> '' then VElt8 := FloatToStr(arrondi(TLA.GetValue(TabElt[8]), 0))
                          else VElt8 := '';
                          EltC9 := TabElt[9];
                          If EltC9 <> '' then VElt9 := FloatToStr(arrondi(TLA.GetValue(TabElt[9]), 0))
                          else VElt9 := '';
                          EltC10 := TabElt[10];
                          If EltC10 <> '' then VElt10 := FloatToStr(arrondi(TLA.GetValue(TabElt[10]), 0))
                          else VElt10 := '';
                          //DEB PT18
                          EltC11 := TabElt[11];
                          If EltC11 <> '' then VElt11 := FloatToStr(arrondi(TLA.GetValue(TabElt[11]), 0))
                          else VElt11 := '';
                          EltC12 := TabElt[12];
                          If EltC12 <> '' then VElt12 := FloatToStr(arrondi(TLA.GetValue(TabElt[12]), 0))
                          else VElt12 := '';
                          //FIN PT18
                        end;
                        EcritureLigne(IdEntreprise, 'S', 13);
                        EcritureLigne(IdEtab, 'S', 14);
                        EcritureLigne(ActiviteMSA, 'I', 4);
                        EcritureLigne(UniteGestion, 'S', 2);
                        EcritureLigne('PE41', 'S', 4);
                        EcritureLigne(SDateDebut, 'S', 8);
                        EcritureLigne(SDateFin, 'S', 8);
                        EcritureLigne(TotalSalaires, 'I', 10);    //PT4
                        EcritureLigne(EltC1, 'I', 2);
                        EcritureLigne(VElt1, 'I', 10);
                        EcritureLigne(EltC2, 'I', 2);
                        EcritureLigne(VElt2, 'I', 10);
                        EcritureLigne(EltC3, 'I', 2);
                        EcritureLigne(VElt3, 'I', 10);
                        EcritureLigne(EltC4, 'I', 2);
                        EcritureLigne(VElt4, 'I', 10);
                        EcritureLigne(EltC5, 'I', 2);
                        EcritureLigne(VElt5, 'I', 10);
                        EcritureLigne(EltC6, 'I', 2);
                        EcritureLigne(VElt6, 'I', 10);
                        EcritureLigne(EltC7, 'I', 2);
                        EcritureLigne(VElt7, 'I', 10);
                        EcritureLigne('', 'S', 53);
                        WriteLN(FEcrt, '');
                        NbEnreg := NbEnreg + 1;
                        If EltC8 <> '' then
                        begin
                             EcritureLigne(IdEntreprise, 'S', 13);
                             EcritureLigne(IdEtab, 'S', 14);
                             EcritureLigne(ActiviteMSA, 'I', 4);
                             EcritureLigne(UniteGestion, 'S', 2);
                             EcritureLigne('PE41', 'S', 4);
                             EcritureLigne(SDateDebut, 'S', 8);
                             EcritureLigne(SDateFin, 'S', 8);
                             EcritureLigne(TotalSalaires, 'I', 10);    //PT4
                             EcritureLigne(EltC8, 'I', 2);
                             EcritureLigne(VElt8, 'I', 10);
                             EcritureLigne(EltC9, 'I', 2);
                             EcritureLigne(VElt9, 'I', 10);
                             EcritureLigne(EltC10, 'I', 2);
                             EcritureLigne(VElt10, 'I', 10);
                             EcritureLigne(EltC11, 'I', 2);  //PT18
                             EcritureLigne(VElt11, 'I', 10); //PT18
                             EcritureLigne(EltC12, 'I', 2);  //PT18
                             EcritureLigne(VElt12, 'I', 10); //PT18
                             EcritureLigne('', 'I', 2);
                             EcritureLigne('', 'I', 10);
                             EcritureLigne('', 'I', 2);
                             EcritureLigne('', 'I', 10);
                             EcritureLigne('', 'S', 53);
                             WriteLN(FEcrt, '');
                             NbEnreg := NbEnreg + 1;
                        end;
                    end;
               end;
        end;
        FreeAndNil(TobUG);
        EltC1 := '';
        VElt1 := '';
        EltC2 := '';
        VElt2 := '';
        EltC3 := '';
        VElt3 := '';
        EltC4 := '';
        VElt4 := '';
        EltC5 := '';
        VElt5 := '';
        //Enregistrement PE41
        if DeclarerAsiette = True then
        begin
          TotalSalaires := FloatToStr(Arrondi(MontantSalairesEtab, 0));
          MontantSalaires := MontantSalaires + MontantSalairesEtab;
          TLA := TobLesAssiettes.FindFirst(['ETABLISSEMENT', 'ACTIVITE'], [Etablissement, ActiviteMSA], False);
          EcritureLigne(IdEntreprise, 'S', 13);
          EcritureLigne(IdEtab, 'S', 14);
          EcritureLigne(ActiviteMSA, 'I', 4);
          EcritureLigne(UniteGestion, 'S', 2);  //PT16
//          if TLA <> nil then EcritureLigne(TLA.GetValue('UG'), 'S', 2)
//          else EcritureLigne('', 'S', 2);
//PT16          EcritureLigne('', 'S', 2);
          EcritureLigne('PE41', 'S', 4);
          EcritureLigne(SDateDebut, 'S', 8);
          EcritureLigne(SDateFin, 'S', 8);
          EcritureLigne(TotalSalaires, 'I', 10);    //PT4
          While TLA <> nil do
          begin
            NumElement := 0;
            For elt := 1 to 10 do TabElt[elt] := '';
            if TLA.GetValue('51') > 0 then
            begin
                 NumElement := NumElement + 1;
                 TabElt[NumElement] := '51';
            end;
            if TLA.GetValue('53') > 0 then
            begin
                 NumElement := NumElement + 1;
                 TabElt[NumElement] := '53';
            end;
            if TLA.GetValue('54') > 0 then
            begin
                 NumElement := NumElement + 1;
                 TabElt[NumElement] := '54';
            end;
            if TLA.GetValue('56') > 0 then
            begin
                 NumElement := NumElement + 1;
                 TabElt[NumElement] := '56';
            end;
            if TLA.GetValue('63') > 0 then
            begin
                 NumElement := NumElement + 1;
                 TabElt[NumElement] := '63';
            end;
            //DEB PT18
            if TLA.GetValue('66') > 0 then
            begin
                 NumElement := NumElement + 1;
                 TabElt[NumElement] := '66';
            end;
            if TLA.GetValue('67') > 0 then
            begin
                 NumElement := NumElement + 1;
                 TabElt[NumElement] := '67';
            end;
            if TLA.GetValue('68') > 0 then
            begin
                 NumElement := NumElement + 1;
                 TabElt[NumElement] := '68';
            end;
            //FIN PT18
            if TLA.GetValue('58') > 0 then
            begin
                 NumElement := NumElement + 1;
                 TabElt[NumElement] := '58';
            end;
            if TLA.GetValue('59') > 0 then
            begin
                 NumElement := NumElement + 1;
                 TabElt[NumElement] := '59';
            end;
            {PT18 if TLA.GetValue('60') > 0 then
            begin
                 NumElement := NumElement + 1;
                 TabElt[NumElement] := '60';
            end;}
            if TLA.GetValue('61') > 0 then
            begin
                 NumElement := NumElement + 1;
                 TabElt[NumElement] := '61';
            end;
            if TLA.GetValue('62') > 0 then
            begin
                 NumElement := NumElement + 1;
                 TabElt[NumElement] := '62';
            end;
            EltC1 := TabElt[1];
            If EltC1 <> '' then VElt1 := FloatToStr(arrondi(TLA.GetValue(TabElt[1]), 0))
            else VElt1 := '';
            EltC2 := TabElt[2];
            If EltC2 <> '' then VElt2 := FloatToStr(arrondi(TLA.GetValue(TabElt[2]), 0))
            else VElt2 := '';
            EltC3 := TabElt[3];
            If EltC3 <> '' then VElt3 := FloatToStr(arrondi(TLA.GetValue(TabElt[3]), 0))
            else VElt3 := '';
            EltC4 := TabElt[4];
            If EltC4 <> '' then VElt4 := FloatToStr(arrondi(TLA.GetValue(TabElt[4]), 0))
            else VElt4 := '';
            EltC5 := TabElt[5];
            If EltC5 <> '' then VElt5 := FloatToStr(arrondi(TLA.GetValue(TabElt[5]), 0))
            else VElt5 := '';
            EltC6 := TabElt[6];
            If EltC6 <> '' then VElt6 := FloatToStr(arrondi(TLA.GetValue(TabElt[6]), 0))
            else VElt6 := '';
            EltC7 := TabElt[7];
            If EltC7 <> '' then VElt7 := FloatToStr(arrondi(TLA.GetValue(TabElt[7]), 0))
            else VElt7 := '';
            EltC8 := TabElt[8];
            If EltC8 <> '' then VElt8 := FloatToStr(arrondi(TLA.GetValue(TabElt[8]), 0))
            else VElt8 := '';
            EltC9 := TabElt[9];
            If EltC9 <> '' then VElt9 := FloatToStr(arrondi(TLA.GetValue(TabElt[9]), 0))
            else VElt9 := '';
            EltC10 := TabElt[10];
            If EltC10 <> '' then VElt10 := FloatToStr(arrondi(TLA.GetValue(TabElt[10]), 0))
            else VElt10 := '';
            //DEB PT18
            EltC11 := TabElt[11];
            If EltC11 <> '' then VElt11 := FloatToStr(arrondi(TLA.GetValue(TabElt[11]), 0))
            else VElt11 := '';
            EltC12 := TabElt[12];
            If EltC12 <> '' then VElt12 := FloatToStr(arrondi(TLA.GetValue(TabElt[12]), 0))
            else VElt12 := '';
            //FIN PT18
            TLA := TobLesAssiettes.FindNext(['ETABLISSEMENT', 'ACTIVITE'], [Etablissement, ActiviteMSA], False);
          end;
          EcritureLigne(EltC1, 'I', 2);
          EcritureLigne(VElt1, 'I', 10);
          EcritureLigne(EltC2, 'I', 2);
          EcritureLigne(VElt2, 'I', 10);
          EcritureLigne(EltC3, 'I', 2);
          EcritureLigne(VElt3, 'I', 10);
          EcritureLigne(EltC4, 'I', 2);
          EcritureLigne(VElt4, 'I', 10);
          EcritureLigne(EltC5, 'I', 2);
          EcritureLigne(VElt5, 'I', 10);
          EcritureLigne(EltC6, 'I', 2);
          EcritureLigne(VElt6, 'I', 10);
          EcritureLigne(EltC7, 'I', 2);
          EcritureLigne(VElt7, 'I', 10);
          EcritureLigne('', 'S', 53);
          WriteLN(FEcrt, '');
          NbEnreg := NbEnreg + 1;
          If EltC8 <> '' then
          begin
               EcritureLigne(IdEntreprise, 'S', 13);
               EcritureLigne(IdEtab, 'S', 14);
               EcritureLigne(ActiviteMSA, 'I', 4);
               EcritureLigne('', 'S', 2);
               EcritureLigne('PE41', 'S', 4);
               EcritureLigne(SDateDebut, 'S', 8);
               EcritureLigne(SDateFin, 'S', 8);
               EcritureLigne(TotalSalaires, 'I', 10);
               EcritureLigne(EltC8, 'I', 2);
               EcritureLigne(VElt8, 'I', 10);
               EcritureLigne(EltC9, 'I', 2);
               EcritureLigne(VElt9, 'I', 10);
               EcritureLigne(EltC10, 'I', 2);
               EcritureLigne(VElt10, 'I', 10);
               EcritureLigne(EltC11, 'I', 2);   //PT18
               EcritureLigne(VElt11, 'I', 10);  //PT18
               EcritureLigne(EltC12, 'I', 2);   //PT18
               EcritureLigne(VElt12, 'I', 10);  //PT18
               EcritureLigne('', 'I', 2);
               EcritureLigne('', 'I', 10);
               EcritureLigne('', 'I', 2);
               EcritureLigne('', 'I', 10);
               EcritureLigne('', 'S', 53);
               WriteLN(FEcrt, '');
               NbEnreg := NbEnreg + 1;
          end;
        end;
      end;
      MoveCurProgressForm(TobEtablissement.Detail[e].GetValue('ET_LIBELLE'));
    end;
    FiniMoveProgressForm;
    FreeAndNil(TobEtablissement);
    FreeAndNil(TobLesAssiettes);
    FreeAndNil(TobActivite);
    //SEGMENT PE11
    Closefile(FEcrt);
    AssignFile(FLect, FileTemp);
    reset(FLect);
    EcritureLignePE11(IdEntreprise, 'S', 13);
    EcritureLignePE11('', 'S', 14);
    EcritureLignePE11('', 'S', 4);
    EcritureLignePE11('', 'S', 2);
    EcritureLignePE11('PE11', 'S', 4);
    EcritureLignePE11(RaisonSoc, 'S', 20);
    EcritureLignePE11('E', 'S', 1);
    EcritureLignePE11(TypeFic, 'S', 1);
    EcritureLignePE11(IntToStr(NbEnreg), 'I', 5);
    EcritureLignePE11('', 'S', 136);
    RecupFichierTemp;
    CloseFile(FZ);
    CloseFile(FLect);
    DeleteFile(PChar(FileTemp));
    CloseFile(FFinal);
    if Erreur = True then
    begin
      DeleteFile(PChar(FileN));
      i := PGIAsk('Il y a des erreurs, #13#10 Voulez vous consulter maintenant le fichier ' + FileZ, Ecran.Caption);
// d PT12
      {$IFDEF EAGLCLIENT}
      if VH_Paie.PGCheminEagl <> '' then FileZ := '"'+VH_Paie.PGCheminEagl + '\ErreurMSA.log"';
      {$ENDIF}
// f PT12
      if i = MrYes then ShellExecute(0, PCHAR('open'), PChar('WordPad'), PChar(FileZ), nil, SW_RESTORE);
    end
    else
    begin
      StDelete := 'SELECT MAX(PES_CHRONOMESS) AS MAXI FROM ENVOISOCIAL';
      QRechDelete := OpenSQL(StDelete, TRUE);
      if not QRechDelete.EOF then Ordre := QRechDelete.FindField('MAXI').AsInteger + 1
      else Ordre := 1;
      Ferme(QRechDelete);
      ChargeTOBENVOI;
      EnregEnvoi.Ordre := Ordre;
      EnregEnvoi.TypeE := 'MSA';
      EnregEnvoi.Millesime := Copy(GetControlText('DATEFIN'), 7, 4);
      ;
      EnregEnvoi.Periodicite := 'T';
      EnregEnvoi.DateD := StrToDate(GetControltext('DATEDEBUT'));
      EnregEnvoi.DateF := StrToDate(GetControltext('DATEFIN'));
      EnregEnvoi.Siret := GetControlText('ESIREN');
      EnregEnvoi.Fraction := '';
      EnregEnvoi.Libelle := Libelle;
      EnregEnvoi.Size := 1;
      EnregEnvoi.NomFic := fichier;
      EnregEnvoi.Statut := '005';
      EnregEnvoi.Monnaie := 'EUR';
      EnregEnvoi.Inst := 'ZCSP';
      EnregEnvoi.EmettSoc := GetControlText('EMETSOC2'); //PT13
      CreeEnvoi(EnregEnvoi);
      LibereTOBENVOI;
      SetControlText('NBSALARIE', IntToStr(NbSalaries));
      SetControlText('MONTANT', FloatToStr(Arrondi(MontantSalaires, 0)));
      Rep := PGIAsk('Voulez-vous visualiser (en vue de l''enregistrer) la déclaration qui sera à joindre avec le fichier afin de l''éditer lors de l''envoi', Ecran.Caption);
      if rep = mrYes then
      begin
        StPages := '';
        {$IFDEF EAGLCLIENT}
        StPages := AglGetCriteres(Pages, FALSE);
        {$ENDIF}
        LanceEtat('E', 'PMS', 'PAJ', True, False, False, Pages, '', '', False, 0, StPages);
      end;
      PGIBox('Le fichier est prêt à être envoyé,#13#10 la prochaine étape est la gestion des envois (module paramètres)', Ecran.caption);
    end;
    CommitTrans;
  except
    Rollback;
    CloseFile(FZ);
    CloseFile(FEcrt);
    CloseFile(FFinal);
  end;
end;

procedure TOF_PGMSAENTREPRISE.EcritureLigne(Valeur, Format: string; Longueur: integer);
var
  i: Integer;
  Heures, Minutes: string;
  Negatif : Boolean;
begin
  if (Format <> 'H') and (Length(Valeur) > Longueur) then Valeur := Copy(Valeur, 1, Longueur);
  if Format = 'I' then
  begin
    Negatif := False;   
    if Valeur = '0' then Valeur := '';
    if Valeur = '' then
    begin
      for i := Length(Valeur) to Longueur - 1 do
      begin
        Valeur := Valeur + ' ';
      end;
    end
    else
    begin
      if Length(Valeur) < Longueur then
      begin
        If StrToFloat(Valeur) < 0 then
        begin
             Valeur := FloatToStr(-StrToFloat(Valeur));
             Negatif := True;
             //If i < longueur - 1 then Valeur := '0' + Valeur
//             else Valeur := '-' + Valeur;
        end;
        for i := Length(Valeur) to Longueur - 1 do
        begin
          begin
               If Negatif then
               begin
                    If i < longueur - 1 then Valeur := '0' + Valeur
                    else Valeur := '-' + Valeur;
               end
               else Valeur := '0' + Valeur;
          end;
        end;
      end;
    end;
  end;
  if Format = 'S' then
  begin
    if Length(Valeur) < Longueur then
    begin
      for i := Length(Valeur) to Longueur - 1 do
      begin
        Valeur := Valeur + ' ';
      end;
    end;
  end;
  //DEBUT PT3
  if Format = 'H' then
  begin
    Heures := ReadTokenPipe(Valeur, ',');
    Minutes := Valeur;
    If Minutes = '' then
    begin
         Valeur := Heures;
         Heures := ReadTokenPipe(Valeur, '.');
         Minutes := Valeur;
    end;
    If Minutes = '' then Minutes := '00';
    Valeur := Heures + Minutes;
    if Length(Valeur) < Longueur then
    begin
      for i := Length(Valeur) to Longueur - 1 do
      begin
        Valeur := '0' + Valeur ;
      end;
    end;
  end;
  //FIN PT3
  Write(FEcrt, Valeur);
end;

procedure TOF_PGMSAENTREPRISE.EcritureLignePE11(Valeur, Format: string; Longueur: integer);
var
  i: Integer;
  Heures, Minutes: string;
begin
  if Valeur = '' then Format := 'S';
  if (Format <> 'H') and (Length(Valeur) > Longueur) then Valeur := Copy(Valeur, 1, Longueur);
  if Format = 'I' then
  begin
    if Valeur = '0' then Valeur := '';
    if Valeur = '' then
    begin
      for i := Length(Valeur) to Longueur - 1 do
      begin
        Valeur := Valeur + ' ';
      end;
    end
    else
    begin
      if Length(Valeur) < Longueur then
      begin
        for i := Length(Valeur) to Longueur - 1 do
        begin
          Valeur := '0' + Valeur;
        end;
      end;
    end;
  end;
  if Format = 'S' then
  begin
    if Length(Valeur) < Longueur then
    begin
      for i := Length(Valeur) to Longueur - 1 do
      begin
        Valeur := Valeur + ' ';
      end;
    end;
  end;
  if Format = 'H' then
  begin
    Heures := ReadTokenPipe(Valeur, ',');
    Minutes := Valeur;
    Valeur := Heures + Minutes;
    if Length(Valeur) < Longueur then
    begin
      for i := Length(Valeur) to Longueur - 1 do
      begin
        Valeur := Valeur + '0';
      end;
    end;
  end;
  Write(FFinal, Valeur);
end;


function TOF_PGMSAENTREPRISE.ControleFichierMSA(Valeur: string; NumSegment: Integer): string;
begin
  Result := '';
  case NumSegment of
    212, 222, 232, 242, 312: // Identifiant établissement
      begin
        if Valeur = '' then Result := 'L''établissement n''est pas renseigné et est obligatoire';
      end;
    213, 223, 233, 243, 313: //Code activité professionelle
      begin
        if Valeur = '' then Result := 'Le code activité professionnelle n''est pas renseigné et est obligatoire';
      end;
    215, 225, 235, 245, 315: // Identifiant salarié = Num SS
      begin
        if Valeur = '' then Result := 'Le n° de sécurité sociale n''est pas renseigné et est obligatoire';
      end;
    216, 226, 236, 246, 316: //Nom
      begin
        if Valeur = '' then Result := 'Le nom n''est pas renseigné et est obligatoire';
      end;
    217, 227, 237, 247, 317: //Prénom
      begin
        if Valeur = '' then Result := 'Le prénom n''est pas renseigné et est obligatoire';
      end;
    218, 228, 238, 248, 318: //Date de naissance
      begin
        if Valeur = '' then Result := 'La date de naissance n''est pas renseigné et est obligatoire';
      end;
    2211: //Code Susp Contrat
      begin
        if (Valeur <> '') and (Valeur <> '01') and (Valeur <> '02') and (Valeur <> '03') and (Valeur <> '04') and (Valeur <> '05') and (Valeur <> '06') then
          Result := 'Le code de suspension de contrat n''est pas valide';
      end;
    2310: //Departement lieu de travail
      begin
        if Valeur <> '' then
        begin
          if not IsNumeric(Valeur) then Result := 'Le département du lieu de travail doit être numérique';
        end;
      end;
    2311: //Commune lieu de travail
      begin
        if Valeur <> '' then
          if not IsNumeric(Valeur) then Result := 'Le code de la commune du lieu de travail doit être numérique';
      end;
    2313: //Code temps plein option temps partiel
      begin
        if Valeur <> '' then
          if (valeur <> '0') and (valeur <> '1') and (valeur <> '2') then Result := 'Le code temps partiel option temps plein n''est pas valide';
      end;
    2314: //Code temps CDD/CDI
      begin
        if Valeur <> '' then
          if (valeur <> 'I') and (valeur <> 'D') then Result := 'Le code CDD/CDI n''est pas valide';
      end;
    2412: //Code poly employeur
      begin
        if Valeur <> '' then
          if (valeur <> '0') and (valeur <> '1') and (valeur <> '2') then Result := 'Le code temps poly-employeur n''est pas valide';
      end;
    2415: //Code classe élèves
      begin
        if Valeur <> '' then
          if (valeur <> '0') and (valeur <> '1') and (valeur <> '2') and (valeur <> '3') and (valeur <> '4') and (valeur <> '5') and (valeur <> '6') then
            Result := 'Le code tclasse pour les élèves d''établissement d''enseignement agricole privé n''est pas valide';
      end;
    3115: //Element rémunératio,
      begin
        if Valeur = '' then Result := 'Aucun élément de rémunération n''est renseigné';
      end;
  end;
end;

function TOF_PGMSAENTREPRISE.ControlePonctuation(var Donnee: string): string;
var
  Longueur, i: Integer;
  Ch: Char;
begin
  Longueur := Length(Donnee);
  for i := 1 to Longueur do
  begin
    Ch := Donnee[i];
    if Ch in [' ', ',', '-', '.', '/', '`', '¨', '°', '´'] then Ch := ' ';
    Donnee[i] := Ch;
  end;
  Result := Donnee;
end;

procedure TOF_PGMSAENTREPRISE.RecupFichierTemp;
var
  S: string;
begin
  while not eof(FLect) do
  begin
    Writeln(FFinal, '');
    Readln(FLect, S);
    Write(FFinal, S);
  end;
end;

procedure TOF_PGMSAENTREPRISE.AffichDeclarant(Sender: TObject);
var St : String;
    Q : TQuery;
begin
if GetControlText('DECLARANT')='' then exit;
//DEBUT PT7
Q := OpenSQL('SELECT PDA_CIVILITE FROM DECLARANTATTEST WHERE PDA_DECLARANTATTES="'+GetControlText('DECLARANT')+'"',True);
If not Q.Eof then SetControlText('CIVILITE',Q.FindField('PDA_CIVILITE').AsString)
else SetControlText('CIVILITE','');
Ferme(Q);
//FINPT7
SetControlText('SIGNATAIRE',RechDom('PGDECLARANTATTEST',GetControlText('DECLARANT'),False));
SetControlText('LIEUEDITION',RechDom('PGDECLARANTVILLE' ,GetControlText('DECLARANT'),False));
//DEBUT PT2
St := RechDom('PGDECLARANTQUAL', GetControlText('DECLARANT'), False);
if St = 'AUT' then SetControlText('QUALITE', RechDom('PGDECLARANTAUTRE', GetControlText('DECLARANT'), False))
else SetControlText('QUALITE', RechDom('PGQUALDECLARANT2', St, False));
//FIN PT2
end;

initialization
  registerclasses([TOF_PGMSAENTREPRISE]);
end.

