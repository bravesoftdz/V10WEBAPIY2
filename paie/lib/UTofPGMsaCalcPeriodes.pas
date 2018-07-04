{***********UNITE*************************************************
Auteur  ...... : JL
Créé le ...... : 24/03/2004
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : MSAMULCALCPER ()
Mots clefs ... : TOF;MSA
*****************************************************************
PT1 14/12/2004 JL V_60 Conersion type variant incorrect si TobPaie.Detail[i].GetValue('HEURESTRAV') = null
PT2 19/05/2005 JL V_60 FQ 12293 Supp. calcul élément 80
PT3 19/05/2005 JL V_60 FQ 12286 Erreur transaction en calcul des périodes
PT4 19/05/2005 JL V_60 FQ 12196 Reprise des salaires à 0
PT5 19/05/2005 JL V_60 FQ 12193 Calcul jours travail
PT6 21/04/2005 V_60 JL Trimestre précédent par défaut
PT7 29/06/2005 JL V_60 Modif calcul nb jours : base ou montant
PT8 26/07/2005 JL V_60 FQ 12439 Correction salariés entréé/sortie même jours
PT9 13/10/2005 JL V_60 FQ 12167 Verification abgrégé tablette activité (num 4 caracteres)
PT10 07/02/2006 JL V_60 FQ 12878 Mise à jour de l'activité dans les évolutions
PT11 04/04/2006 JL V_65 Gestion UG MSA
PT12 18/07/2007 JL V_80 FQ 14481 Gestion des arrêts maladie pour évolution
PT13 31/07/2008 JL V_80 FQ 13816 Ajout calcul élément de rémunération
PT14 12/11/2007 FC V_80 FQ 14892 Evolution cahier des charges Octobre 2007
PT14-1 10/12/2007 FC V_80 Réalignement entre la sélection par ligne et la sélection globale
PT16 12/03/2008 FC V_80 FQ 15303 2 rubriques d'heures TEPA sur le taux de majoration 25
PT17 12/03/2008 FC V_80 FQ 15307 msaperiodePE31 : type element calcule 13
PT18 28/03/2008 FC V_80 FQ 15333 Corrections MSA cahier des charges
PT19 04/04/2008 FC V_80 Suppression préalable des enregistrements car dans le cas de calculs successifs
                        la remise à 0 de champs via une tob ne fonctionne pas. Pour lui pas de modification
                        donc l'ancienne valeur reste (valeurs d'initialisation)
PT20 20/05/2008 FC V_80 FQ 15395 Si cumul='44' alors toujours mettre 30 dans type de rémunération
PT21 20/05/2008 FC V_80 FQ 15396 Suppression code calcul 82
PT22 21/05/2008 FC V_80 FQ 15394 Rajout code calcul 25
PT23 23/06/2008 FC V_810 FQ 15488 Faute d'orthographe
}
unit UTofPGMsaCalcPeriodes;

interface

uses
  {$IFDEF VER150}
  Variants,
  {$ENDIF}
  StdCtrls, Controls, Classes,
  {$IFNDEF EAGLCLIENT}
  db, {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF} Mul, Fe_Main, HDB, DBGrids,
  {$ELSE}
  eMul, MainEAGL,
  {$ENDIF}
  forms, sysutils, ComCtrls, HCtrls, HEnt1, HMsgBox, UTOF, UTob, HTB97, PG_OutilsEnvoi, HStatus,
  EntPaie, AGLInit, ed_tools, ParamDat, HQry, ShellAPI, windows, PGOutils,PGOutils2,ParamSoc;

type
  TOF_PGMSACALCPERIODES = class(TOF)
    procedure OnUpdate; override;
    procedure OnArgument(S: string); override;
  private
    procedure LancerTraitement(Sender: Tobject);
    procedure MajElementsSalarie;
    procedure CalculPeriode;
    procedure DateElipsisClick(Sender: TObject);
    procedure ExitEdit(Sender: TObject);
    Function ControlActivite : Boolean; //PT9
  end;

implementation

procedure TOF_PGMSACALCPERIODES.OnUpdate;
begin
  inherited;
end;

procedure TOF_PGMSACALCPERIODES.OnArgument(S: string);
var
  DateDebut, DateFin: TDateTime;
  BCalcul: TToolBarButton97;
  Arg: string;
  aa, mm, jj, Mois: word;
  THDate,Edit : THEdit;
begin
  inherited;
  Arg := ReadTokenPipe(S, ';');
  BCalcul := TToolBarButton97(GetControl('BCALCULER'));
  if BCalcul <> nil then BCalcul.OnClick := LancerTraitement;
  DecodeDate(V_PGI.DateEntree, aa, mm, jj);
  Mois := 1;
  if (mm < 7) and (mm > 3) then Mois := 4;
  if (mm < 10) and (mm > 6) then Mois := 7;
  if mm > 9 then Mois := 10;
   DateDebut := EncodeDate(aa, Mois, 1);
  DateDebut := PlusMois(DateDebut,-3); //PT6
  DateFin := PlusMois(DateDebut, 2);
  DateFin := FinDeMois(DateFin);
  SetControlText('DATEDEBUT', DateToStr(Datedebut));
  SetControltext('DATEFIN', dateToStr(DateFin));
  THDate := THEdit(GetControl('DATEDEBUT'));
  if THDate <> nil then THDate.OnElipsisClick := DateElipsisClick;
  THDate := THEdit(GetControl('DATEFIN'));
  if THDate <> nil then THDate.OnElipsisClick := DateElipsisClick;
  SetControlText('TYPEMAJ', 'PER');
  Edit := THEdit(GetControl('PSA_SALARIE'));
  if Edit <> nil then Edit.OnExit := ExitEdit;
end;

procedure TOF_PGMSACALCPERIODES.LancerTraitement(Sender: Tobject);
begin
   if Not ControlActivite then exit; //PT9
  if GetControlText('TYPEMAJ') = 'PER' then CalculPeriode
  else MajElementsSalarie;
end;

procedure TOF_PGMSACALCPERIODES.MajElementsSalarie;
var
  TobEvolutions, TobPeriodes, TobSalaries, TS: Tob;
//  Q: TQuery;
  {$IFNDEF EAGLCLIENT}
  Liste: THDBGrid;
  {$ELSE}
  Liste: THGrid;
  {$ENDIF}
  Q_Mul: THQuery;
  Salarie, Nom, Prenom, NumSS, Activite,UGMSA: string;
  dateDebut, DateFin, DateNaiss: TDateTime;
  i, j: Integer;
  St : String;
begin
  DateDebut := StrToDate(GetControlText('DATEDEBUT'));
  DateFin := StrToDate(GetControlText('DATEFIN'));
  {$IFNDEF EAGLCLIENT}
  Liste := THDBGrid(GetControl('FLISTE'));
  {$ELSE}
  Liste := THGrid(GetControl('FLISTE'));
  {$ENDIF}
  if Liste = nil then Exit;
  Q_Mul := THQuery(Ecran.FindComponent('Q'));
  if Q_Mul = nil then Exit;
  if (Liste.NbSelected = 0) and (TFMul(Ecran).BSelectAll.Down = False) then
  begin
    PGIBox('Aucun élément sélectionné', Ecran.Caption);
    Exit;
  end;
  {$IFDEF EAGLCLIENT}
  if (TFMul(Ecran).bSelectAll.Down) then TFMul(Ecran).Fetchlestous;
  {$ENDIF}
//Optimisation  Q := OpenSQL('SELECT PSA_SALARIE,PSA_LIBELLE,PSA_PRENOM,PSA_NUMEROSS,PSA_DATENAISSANCE,PSE_MSAACTIVITE,PSE_MSAUNITEGES FROM SALARIES' +
//    ' LEFT JOIN DEPORTSAL ON PSA_SALARIE=PSE_SALARIE', True);
//  TobSalaries.LoadDetailDB('LesSalaries', '', '', Q, False);
  St := 'SELECT PSA_SALARIE,PSA_LIBELLE,PSA_PRENOM,PSA_NUMEROSS,PSA_DATENAISSANCE,PSE_MSAACTIVITE,PSE_MSAUNITEGES FROM SALARIES' +
    ' LEFT JOIN DEPORTSAL ON PSA_SALARIE=PSE_SALARIE';
  TobSalaries := Tob.Create('LesSalaries', nil, -1);
  TobSalaries.LoadDetailDBFromSQL ('LesSalaries', St);
//  Ferme(Q);
  if ((Liste.nbSelected) > 0) and (not Liste.AllSelected) then
  begin
    InitMoveProgressForm(nil, 'Début du traitement', 'Veuillez patienter SVP ...', Liste.nbSelected, FALSE, TRUE);
    InitMove(Liste.nbSelected, '');
    for i := 0 to Liste.NbSelected - 1 do
    begin
      Liste.GotoLeBOOKMARK(i);
      {$IFDEF EAGLCLIENT}
      TFmul(Ecran).Q.TQ.Seek(Liste.Row - 1);
      {$ENDIF}
      Salarie := TFmul(Ecran).Q.FindField('PSE_SALARIE').asstring;
      TS := TobSalaries.FindFirst(['PSA_SALARIE'], [Salarie], False);
      if TS <> nil then
      begin
        Nom := TS.GetValue('PSA_LIBELLE');
        Prenom := TS.GetValue('PSA_PRENOM');
        DateNaiss := TS.GetValue('PSA_DATENAISSANCE');
        NumSS := TS.GetValue('PSA_NUMEROSS');
        Activite := TS.GetValue('PSE_MSAACTIVITE');
        UGMSA := TS.GetValue('PSE_MSAUNITEGES');
//Optimisation        Q := OpenSQL('SELECT * FROM MSAPERIODESPE31 WHERE PE3_SALARIE="' + Salarie + '" ' +
//          'AND PE3_DATEDEBUT>="' + UsDateTime(DateDebut) + '" AND PE3_DATEFIN<="' + UsDateTime(DateFin) + '"', True);
//        TobPeriodes.LoadDetailDB('MSAPERIODESPE31', '', '', Q, False);
        St := 'SELECT * FROM MSAPERIODESPE31 WHERE PE3_SALARIE="' + Salarie + '" ' +
          'AND PE3_DATEDEBUT>="' + UsDateTime(DateDebut) + '" AND PE3_DATEFIN<="' + UsDateTime(DateFin) + '"';
        TobPeriodes := Tob.Create('MSAPERIODESPE31', nil, -1);
        TobPeriodes.LoadDetailDBFromSQL ('MSAPERIODESPE31', St);
//        Ferme(Q);
        for j := 0 to TobPeriodes.Detail.Count - 1 do
        begin
          TobPeriodes.Detail[j].PutValue('PE3_NOM', Nom);
          TobPeriodes.Detail[j].PutValue('PE3_PRENOM', Prenom);
          TobPeriodes.Detail[j].PutValue('PE3_NUMEROSS', NumSS);
          TobPeriodes.Detail[j].PutValue('PE3_DATENAISSANCE', DateNaiss);
          TobPeriodes.Detail[j].PutValue('PE3_MSAACTIVITE', Activite);
          TobPeriodes.Detail[j].PutValue('PE3_UGMSA', UgMSA);//PT11
          TobPeriodes.Detail[j].UpdateDB(False);
        end;
        FreeAndNil(TobPeriodes);
        //DEBUT PT10
//Optimisation        Q := OpenSQL('SELECT * FROM MSAEVOLUTIONSPE2 WHERE PE2_SALARIE="' + Salarie + '" ' +
//          'AND PE2_DATEEFFET>="' + UsDateTime(DateDebut) + '" AND PE2_DATEEFFET<="' + UsDateTime(DateFin) + '"', True);
//        TobEvolutions.LoadDetailDB('MSAEVOLUTIONSPE2', '', '', Q, False);
        St := 'SELECT * FROM MSAEVOLUTIONSPE2 WHERE PE2_SALARIE="' + Salarie + '" ' +
          'AND PE2_DATEEFFET>="' + UsDateTime(DateDebut) + '" AND PE2_DATEEFFET<="' + UsDateTime(DateFin) + '"';
        TobEvolutions := Tob.Create('MSAEVOLUTIONSPE2', nil, -1);
        TobEvolutions.LoadDetailDBFromSQL ('MSAEVOLUTIONSPE2', St);
//        Ferme(Q);
        for j := 0 to TobEvolutions.Detail.Count - 1 do
        begin
          TobEvolutions.Detail[j].PutValue('PE2_NOM', Nom);
          TobEvolutions.Detail[j].PutValue('PE2_PRENOM', Prenom);
          TobEvolutions.Detail[j].PutValue('PE2_NUMEROSS', NumSS);
          TobEvolutions.Detail[j].PutValue('PE2_DATENAISSANCE', DateNaiss);
          TobEvolutions.Detail[j].PutValue('PE2_MSAACTIVITE', Activite);
          TobEvolutions.Detail[j].UpdateDB(False);
        end;
        FreeAndNil(TobEvolutions);
        //FIN PT10
      end;
      MoveCurProgressForm(Nom);
    end;
    FiniMoveProgressForm;
  end
  else if liste.AllSelected then
  begin
    InitMoveProgressForm(nil, 'Début du traitement', 'Veuillez patienter SVP ...', TFmul(Ecran).Q.RecordCount, FALSE, TRUE);
    InitMove(TFmul(Ecran).Q.RecordCount, '');
    Q_Mul.First;
    while not Q_Mul.EOF do
    begin
      Salarie := TFmul(Ecran).Q.FindField('PSE_SALARIE').asstring;
      TS := TobSalaries.FindFirst(['PSA_SALARIE'], [Salarie], False);
      if TS <> nil then
      begin
        Nom := TS.GetValue('PSA_LIBELLE');
        Prenom := TS.GetValue('PSA_PRENOM');
        DateNaiss := TS.GetValue('PSA_DATENAISSANCE');
        NumSS := TS.GetValue('PSA_NUMEROSS');
        Activite := TS.GetValue('PSE_MSAACTIVITE');
        UGMSA := TS.GetValue('PSE_MSAUNITEGES');
//Optimisation        Q := OpenSQL('SELECT * FROM MSAPERIODESPE31 WHERE PE3_SALARIE="' + Salarie + '" ' +
//          'AND PE3_DATEDEBUT>="' + UsDateTime(DateDebut) + '" AND PE3_DATEFIN<="' + UsDateTime(DateFin) + '"', True);
//        TobPeriodes.LoadDetailDB('MSAPERIODESPE31', '', '', Q, False);
        St := 'SELECT * FROM MSAPERIODESPE31 WHERE PE3_SALARIE="' + Salarie + '" ' +
          'AND PE3_DATEDEBUT>="' + UsDateTime(DateDebut) + '" AND PE3_DATEFIN<="' + UsDateTime(DateFin) + '"';
        TobPeriodes := Tob.Create('MSAPERIODESPE31', nil, -1);
        TobPeriodes.LoadDetailDBFromSQL ('MSAPERIODESPE31', St);
//        Ferme(Q);
        for j := 0 to TobPeriodes.Detail.Count - 1 do
        begin
          TobPeriodes.Detail[j].PutValue('PE3_NOM', Nom);
          TobPeriodes.Detail[j].PutValue('PE3_PRENOM', Prenom);
          TobPeriodes.Detail[j].PutValue('PE3_NUMEROSS', NumSS);
          TobPeriodes.Detail[j].PutValue('PE3_DATENAISSANCE', DateNaiss);
          TobPeriodes.Detail[j].PutValue('PE3_MSAACTIVITE', Activite);
          TobPeriodes.Detail[j].PutValue('PE3_UGMSA', UgMSA);//PT11          
          TobPeriodes.Detail[j].UpdateDB(False);
        end;
        FreeAndNil(TobPeriodes);
//Optimisation        Q := OpenSQL('SELECT * FROM MSAEVOLUTIONSPE2 WHERE PE2_SALARIE="' + Salarie + '" ' +
//          'AND PE2_DATEEFFET>="' + UsDateTime(DateDebut) + '" AND PE2_DATEEFFET<="' + UsDateTime(DateFin) + '"', True);
//        TobEvolutions.LoadDetailDB('MSAEVOLUTIONSPE2', '', '', Q, False);
        St := 'SELECT * FROM MSAEVOLUTIONSPE2 WHERE PE2_SALARIE="' + Salarie + '" ' +
          'AND PE2_DATEEFFET>="' + UsDateTime(DateDebut) + '" AND PE2_DATEEFFET<="' + UsDateTime(DateFin) + '"';
        TobEvolutions := Tob.Create('MSAEVOLUTIONSPE2', nil, -1);
        TobEvolutions.LoadDetailDBFromSQL ('MSAEVOLUTIONSPE2', St);
//        Ferme(Q);
        for j := 0 to TobEvolutions.Detail.Count - 1 do
        begin
          TobEvolutions.Detail[j].PutValue('PE2_NOM', Nom);
          TobEvolutions.Detail[j].PutValue('PE2_PRENOM', Prenom);
          TobEvolutions.Detail[j].PutValue('PE2_NUMEROSS', NumSS);
          TobEvolutions.Detail[j].PutValue('PE2_DATENAISSANCE', DateNaiss);
          TobEvolutions.Detail[j].PutValue('PE2_MSAACTIVITE', Activite);
          TobEvolutions.Detail[j].UpdateDB(False);
        end;
        FreeAndNil(TobEvolutions);
      end;
      Q_Mul.Next;
    end;
    FiniMoveProgressForm;
  end;
  FreeAndNil(TobSalaries); //PT14
end;

procedure TOF_PGMSACALCPERIODES.CalculPeriode;
var
  TobPaie, TobSalaries, TS, TobMSAPeriodes, TC, TobEvolutions, TE,TobAbs : Tob;
  Q: TQuery;
  Salarie, Nom, Prenom, NomFicRapport, EtablissementS, NumSS: string;
  ActiviteMsa, Sexe: string;
  DateDebut, dateFin, DateNaiss, DebPaie, FinPaie: TDateTime;
  Montant01, Montant02, Montant09, Montant12, Montant13, Montant14,  Montant81, Montant82, Montant83,MontantR5850: Double;  //PT17
  {$IFNDEF EAGLCLIENT}
  Liste: THDBGrid;
  {$ELSE}
  Liste: THGrid;
  {$ENDIF}
  Q_Mul: THQuery;
  FRapport: TextFile;
  NbAbs,i, j, NumElement, x,r: Integer;
  MajEvolutions: Boolean;
  DateEntree, DateSortie: TDateTime;
  PlafondSS: Double;
  NbHeures,Jours: Double;
  RubJours,UGMsa : String;
  Typecal,TypeBase,TypeMontant,ChampBul : String;
  Cumul : String;
  NbHeuresSup : double;
  St : String;
begin
//  UGMsa := GetParamSoc('SO_PGMSAUNITEG'); //PT5  PT11 mis en commentaire
  RubJours := GetParamSoc('SO_PGMSARUBJOURS'); //PT5
  // DEBUT PT7
  ChampBul := 'PHB_MTREM';
  Typecal := '';
  TypeBase := '';
  TypeMontant := '';
  Q := OpenSQL('SELECT PRM_CODECALCUL,PRM_TYPEBASE,PRM_TYPEMONTANT FROM REMUNERATION WHERE ##PRM_PREDEFINI## PRM_RUBRIQUE="'+RubJours+'"',True);
  If Not Q.Eof then
  begin
        Typecal := Q.FindField('PRM_CODECALCUL').AsString;
        TypeBase := Q.FindField('PRM_TYPEBASE').AsString;
        TypeMontant := Q.FindField('PRM_TYPEMONTANT').AsString;
  end;
  Ferme(Q);
  If TypeCal = '01' then
  begin
       If typeBase = '00' then ChampBul := 'PHB_BASEREM';
  end
  else ChampBul := 'PHB_BASEREM';
  //FIN PT7
  {$IFNDEF EAGLCLIENT}
  Liste := THDBGrid(GetControl('FLISTE'));
  {$ELSE}
  Liste := THGrid(GetControl('FLISTE'));
  {$ENDIF}
  if Liste = nil then Exit;
  Q_Mul := THQuery(Ecran.FindComponent('Q'));
  if Q_Mul = nil then Exit;
  if (Liste.NbSelected = 0) and (TFMul(Ecran).BSelectAll.Down = False) then
  begin
    PGIBox('Aucun élément sélectionné', Ecran.Caption);
    Exit;
  end;
  MajEvolutions := False;
  case PGIAskCancel('Le calcul va être effectué pour la période du ' + GetControlText('DATEDEBUT') + ' au ' + GetControlText('DATEFIN') +
    '#13#10 Voulez-vous également mettre à jour les évolutions du salarié pour cette période ?', Ecran.Caption) of
    mrYes: MajEvolutions := True;
    mrCancel: Exit
  end;
  Q := OpenSQL('SELECT PEL_MONTANTEURO FROM ELTNATIONAUX WHERE ##PEL_PREDEFINI## PEL_CODEELT="0001" ' +
    'AND PEL_DATEVALIDITE<="' + UsDateTime(Date) + '" ORDER BY PEL_DATEVALIDITE DESC', true);
  if not Q.Eof then PlafondSS := Q.Fields[0].AsFloat
  else PlafondSS := 0;
  Ferme(Q);
  DateDebut := StrToDate(GetControlText('DATEDEBUT'));
  DateFin := StrToDate(GetControlText('DATEFIN'));
  {$IFDEF EAGLCLIENT}
  NomFicRapport := VH_PAIE.PGCheminEagl + '\' + 'MSA_PGI.log';
  {$ELSE}
  NomFicRapport := V_PGI.DatPath + '\' + 'MSA_PGI.log';
  {$ENDIF}
  if FileExists(NomFicRapport) then DeleteFile(PChar(NomFicRapport));
  AssignFile(FRapport, NomFicRapport);
  {$I-}ReWrite(FRapport);
  {$I+}
  if IoResult <> 0 then
  begin
    PGIBox('Fichier inaccessible : ' + NomFicRapport,
      'Rapport calcul des périodes déclarations MSA');
  end;
  try
    begintrans;
    Writeln(FRapport, '');
    Writeln(FRapport, 'Début du calcul des traitements et salaires MSA : ' + DateTimeToStr(Now));
    {$IFDEF EAGLCLIENT}
    if (TFMul(Ecran).bSelectAll.Down) then
      TFMul(Ecran).Fetchlestous;
    {$ENDIF}
    St := 'SELECT PSA_SALARIE,PSA_LIBELLE,PSA_PRENOM,PSA_DATENAISSANCE,PSE_MSAACTIVITE,PSA_NUMEROSS,PSA_SEXE,PSE_MSAUNITEGES' +
      ',PSA_DATEENTREE,PSA_DATESORTIE,PSA_ETABLISSEMENT,PSA_CONDEMPLOI,PSA_SALAIRETHEO,PSA_MULTIEMPLOY,PSA_SALAIREMULTI FROM SALARIES' +
      ' LEFT JOIN DEPORTSAL ON PSA_SALARIE=PSE_SALARIE';
    TobSalaries := Tob.Create('LesSalaries', nil, -1);
    TobSalaries.LoadDetailDBFromSQL ('TobSalaries', St);
    if ((Liste.nbSelected) > 0) and (not Liste.AllSelected) then
    begin
      InitMoveProgressForm(nil, 'Début du traitement', 'Veuillez patienter SVP ...', Liste.nbSelected, FALSE, TRUE);
      InitMove(Liste.nbSelected, '');
      for j := 0 to Liste.NbSelected - 1 do
      begin
        Liste.GotoLeBOOKMARK(j);
        {$IFDEF EAGLCLIENT}
        TFmul(Ecran).Q.TQ.Seek(Liste.Row - 1);
        {$ENDIF}
        Salarie := TFmul(Ecran).Q.FindField('PSE_SALARIE').asstring;
        TS := TobSalaries.FindFirst(['PSA_SALARIE'], [Salarie], False);
        EtablissementS := TS.GetValue('PSA_ETABLISSEMENT');
        Nom := TS.GetValue('PSA_LIBELLE');
        Prenom := TS.GetValue('PSA_PRENOM');
        Writeln(FRapport, '');
        Writeln(FRapport, '     Salarie :' + Salarie + ' ' + Nom + ' ' + Prenom);
        DateNaiss := TS.GetValue('PSA_DATENAISSANCE');
        NumSS := TS.GetValue('PSA_NUMEROSS');
        Sexe := TS.GetValue('PSA_SEXE');
        ActiviteMsa := TS.Getvalue('PSE_MSAACTIVITE');
        UGMsa := TS.Getvalue('PSE_MSAUNITEGES');
        if TestNumeroSS(NumSS, Sexe) < 0 then Writeln(FRapport, '- véifier le N° de sécurité sociale');
        if DateNaiss = IDate1900 then Writeln(FRapport, '- la date de naissance n''est pas renseignée');
        if ActiviteMsa = '' then Writeln(FRapport, '- l''activité n''est pas renseignée');
        St := 'SELECT PPU_SALARIE,PPU_ETABLISSEMENT,PPU_DATEDEBUT,PPU_DATEFIN,PPU_PAYELE,PPU_CBRUT,(PHC_MONTANT) AS HEURESTRAV' +
          ' FROM PAIEENCOURS' +
          ' LEFT JOIN HISTOCUMSAL ON PPU_DATEDEBUT=PHC_DATEDEBUT AND PPU_DATEFIN=PHC_DATEFIN AND PPU_SALARIE=PHC_SALARIE AND PHC_CUMULPAIE="21"' +
          ' WHERE PPU_SALARIE="' + Salarie + '" AND' +
          ' PPU_DATEDEBUT>="' + UsDatetime(DateDebut) + '"' +
          ' AND PPU_DATEFIN<="' + UsDatetime(DateFin) + '"';
        TobPaie := Tob.Create('Les paies', nil, -1);
        TobPaie.LoadDetailDBFromSQL ('PAIEENCOURS', St);

        //DEB PT19
        //Supprimer les enregistrements déjà présents car la remise à 0 d'un champ par une tob
        //ne fonctionne pas
        St := 'DELETE FROM MSAPERIODESPE31' +
          ' WHERE PE3_SALARIE="' + Salarie + '"' +
          ' AND PE3_DATEDEBUT>="' + USDATETIME(DateDebut) + '"' +
          ' AND PE3_DATEFIN<="' + USDATETIME(DateFin) + '"';
        ExecuteSQL(St);
        //FIN PT19

        St := 'SELECT * FROM MSAPERIODESPE31 WHERE PE3_SALARIE="' + Salarie + '"' +
          ' AND PE3_DATEDEBUT>="' + UsDatetime(DateDebut) + '"' +
          ' AND PE3_DATEFIN<="' + UsDatetime(DateFin) + '"';
        TobMSAPeriodes := Tob.create('MSAPERIODESPE31', nil, -1);
        TobMSAPeriodes.LoadDetailDBFromSQL ('MSAPERIODESPE31', St);
        if TobPaie.Detail.Count <= 0 then Writeln(FRapport, '- Ce salarié n''a aucune période de paie');
        for i := 0 to TobPaie.Detail.Count - 1 do
        begin
          DebPaie := TobPaie.Detail[i].GetValue('PPU_DATEDEBUT');
          FinPaie := TobPaie.Detail[i].GetValue('PPU_DATEFIN');
          If TobPaie.Detail[i].GetValue('HEURESTRAV') <> null then
          begin
               if IsNumeric(TobPaie.Detail[i].GetValue('HEURESTRAV')) then NbHeures := TobPaie.Detail[i].GetValue('HEURESTRAV')  // PT1 et PT3
               else NbHeures := 0;
          end
          else NbHeures := 0;
          if NbHeures = 0 then Writeln(FRapport, '- Le nombre d''heures de la période du ' + DateToStr(debPaie) + ' au ' + DateToStr(FinPaie) + ' est égale à 0');
          //                                        Montant01 := TobPaie.Detail[i].GetValue('PPU_CBRUT');
          Montant01 := 0;
          Montant02 := 0;
          Montant09 := 0;
          Montant12 := 0;
          Montant13 := 0;
          Montant14 := 0;
          Montant81 := 0;
          Montant82 := 0;
          Montant83 := 0;
          MontantR5850 := 0;  //PT17
          //DEBUT PT5
          Q := OpenSQL('SELECT SUM('+ChampBul+') MONTANT FROM HISTOBULLETIN WHERE PHB_SALARIE="'+Salarie+'" '+
          'AND PHB_NATURERUB="AAA" AND PHB_RUBRIQUE="'+RubJours+'" '+
          'AND PHB_DATEDEBUT>="'+UsDateTime(DebPaie)+'" '+
          'AND PHB_DATEFIN<="'+UsDateTime(FinPaie)+'"',True);
          If Not Q.Eof then Jours := Q.FindField('MONTANT').AsFloat
          else Jours := 0;
          Ferme(Q);
          //FIN PT5
          TC := Tob.Create('MSAPERIODESPE31', TobMSAPeriodes, -1);
          TC.PutValue('PE3_SALARIE', Salarie);
          TC.PutValue('PE3_NOM', Nom);
          TC.PutValue('PE3_PRENOM', Prenom);
          TC.PutValue('PE3_DATENAISSANCE', DateNaiss);
          TC.PutValue('PE3_MSAACTIVITE', ActiviteMsa);
          TC.PutValue('PE3_NUMEROSS', NumSS);
          TC.PutValue('PE3_ETABLISSEMENT', TobPaie.Detail[i].GetValue('PPU_ETABLISSEMENT'));
          TC.PutValue('PE3_DATEDEBUT', TobPaie.Detail[i].GetValue('PPU_DATEDEBUT'));
          TC.PutValue('PE3_DATEFIN', TobPaie.Detail[i].GetValue('PPU_DATEFIN'));
          TC.PutValue('PE3_ORDRE', 0);
          TC.PutValue('PE3_UGMSA', UGMsa);
          TC.PutValue('PE3_TOPPLAF', '-');
          TC.PutValue('PE3_PEXOMSA', '');
          TC.PutValue('PE3_NBJOURS', Jours);
          TC.PutValue('PE3_NBHEURES', NbHeures);
          For r := 1 to 5 do
          begin
            TC.PutValue('PE3_ELTCALCMSA' + IntToStr(r), '');
            TC.PutValue('PE3_MONTANTELT' + IntToStr(r), 0);
          end;
          Q := OpenSQL('SELECT (PHC_MONTANT) MONTANT' +
            ' FROM HISTOCUMSAL' +
            ' WHERE PHC_SALARIE="' + Salarie + '" AND' +
            ' PHC_CUMULPAIE="02" AND' +
            ' PHC_DATEDEBUT>="' + UsDatetime(DebPaie) + '"' +
            ' AND PHC_DATEFIN<="' + UsDatetime(FinPaie) + '"', True);
          if not Q.Eof then Montant01 := Q.FindField('MONTANT').AsFloat;
          Ferme(Q);
          Q := OpenSQL('SELECT SUM(PHB_MTREM) AS MTPREAVIS' +
            ' FROM HISTOBULLETIN' +
            ' LEFT JOIN REMUNERATION ON' +
            ' PHB_NATURERUB=PRM_NATURERUB AND' +
            ' ##PRM_PREDEFINI## PHB_RUBRIQUE=PRM_RUBRIQUE WHERE' +
            ' PRM_INDEMPREAVIS="X" AND' +
            ' PHB_SALARIE="' + Salarie + '" AND' +
            ' PHB_DATEFIN>="' + UsDateTime(DebPaie) + '" AND' +
            ' PHB_DATEFIN<="' + UsDateTime(FinPaie) + '"', True);
          if not Q.Eof then Montant02 := Q.FindField('MTPREAVIS').AsFloat;
          Ferme(Q);
          Q := OpenSQL('SELECT SUM(PHB_MTREM) AS MTCP' +
            ' FROM HISTOBULLETIN' +
            ' LEFT JOIN REMUNERATION ON' +
            ' PHB_NATURERUB=PRM_NATURERUB AND' +
            ' ##PRM_PREDEFINI## PHB_RUBRIQUE=PRM_RUBRIQUE WHERE' +
            ' PRM_INDEMCOMPCP="X" AND' +
            ' PHB_SALARIE="' + Salarie + '" AND' +
            ' PHB_DATEFIN>="' + UsDateTime(DebPaie) + '" AND' +
            ' PHB_DATEFIN<="' + UsDateTime(FinPaie) + '"', True);
          if not Q.Eof then Montant09 := Q.FindField('MTCP').AsFloat;
          Ferme(Q);
          //DEBUT PT13
          Q := OpenSQL('SELECT SUM(PHB_MTREM) AS MTCP' +
            ' FROM HISTOBULLETIN' +
            ' LEFT JOIN REMUNERATION ON' +
            ' PHB_NATURERUB=PRM_NATURERUB AND' +
            ' ##PRM_PREDEFINI## PHB_RUBRIQUE=PRM_RUBRIQUE WHERE' +
            ' PRM_RUBRIQUE="6000" AND' +
            ' PHB_SALARIE="' + Salarie + '" AND' +
            ' PHB_DATEFIN>="' + UsDateTime(DebPaie) + '" AND' +
            ' PHB_DATEFIN<="' + UsDateTime(FinPaie) + '"', True);
          if not Q.Eof then Montant12 := Q.FindField('MTCP').AsFloat;
          Ferme(Q);
          //FIN PT13
          Q := OpenSQL('SELECT (PHC_MONTANT) MONTANT' +
            ' FROM HISTOCUMSAL' +
            ' WHERE PHC_SALARIE="' + Salarie + '" AND' +
            ' PHC_CUMULPAIE="05" AND' +
            ' PHC_DATEDEBUT>="' + UsDatetime(DebPaie) + '"' +
            ' AND PHC_DATEFIN<="' + UsDatetime(FinPaie) + '"', True);
          if not Q.Eof then Montant13 := Q.FindField('MONTANT').AsFloat;
          Ferme(Q);

{PT18          //DEB PT17
          Q := OpenSQL('SELECT PHB_MTPATRONAL FROM HISTOBULLETIN' +
            ' WHERE PHB_SALARIE="' + Salarie + '" AND ' +
            ' PHB_NATURERUB="COT" AND PHB_RUBRIQUE="5850" AND ' +
            ' PHB_DATEDEBUT>="' + UsDatetime(DebPaie) + '"' +
            ' AND PHB_DATEFIN<="' + UsDatetime(FinPaie) + '"', True);
          if not Q.Eof then MontantR5850 := Q.FindField('PHB_MTPATRONAL').AsFloat;
          Ferme(Q);
          //FIN PT17
}
//PT18          if (Montant13 <> 0) and ((Montant13 = Montant01) or (MontantR5850 >= 0)) then Montant13 := 0;  //PT17
          if (Montant13 = Montant01) then Montant13 := 0;  //PT18

          Q := OpenSQL('SELECT (PHC_MONTANT) MONTANT' +
            ' FROM HISTOCUMSAL' +
            ' WHERE PHC_SALARIE="' + Salarie + '" AND' +
            ' PHC_CUMULPAIE="39" AND' +
            ' PHC_DATEDEBUT>="' + UsDatetime(DebPaie) + '"' +
            ' AND PHC_DATEFIN<="' + UsDatetime(FinPaie) + '"', True);
          if not Q.Eof then Montant14 := Q.FindField('MONTANT').AsFloat;
          Ferme(Q);
          if TS.GetValue('PSA_CONDEMPLOI') = 'P' then
          begin
            Montant81 := TS.GetValue('PSA_SALAIRETHEO');
//PT21            Montant82 := Montant13;
//PT18            Montant13 := 0;
          end;
          Montant13 := 0;  //PT18
          if TS.GetValue('PSA_MULTIEMPLOY') = 'X' then
          begin
            Montant83 := TS.GetValue('PSA_SALAIREMULTI');
          end;
          Montant01 := Montant01 - Montant02 - Montant09 - Montant14;
          NumElement := 0;
//          if Montant01 <> 0 then      PT4
//          begin
            NumElement := NumElement + 1;
            TC.PutValue('PE3_ELTCALCMSA' + IntToStr(NumElement), '01');
            TC.PutValue('PE3_MONTANTELT' + IntToStr(NumElement), Arrondi(Montant01, 0));
//          end;
          if Montant02 <> 0 then
          begin
            NumElement := NumElement + 1;
            TC.PutValue('PE3_ELTCALCMSA' + IntToStr(NumElement), '02');
            TC.PutValue('PE3_MONTANTELT' + IntToStr(NumElement), Arrondi(Montant02, 0));
          end;
          if Montant09 <> 0 then
          begin
            NumElement := NumElement + 1;
            TC.PutValue('PE3_ELTCALCMSA' + IntToStr(NumElement), '09');
            TC.PutValue('PE3_MONTANTELT' + IntToStr(NumElement), Arrondi(Montant09, 0));
          end;
          //DEBUT PT13
          if Montant12 <> 0 then
          begin
            NumElement := NumElement + 1;
            TC.PutValue('PE3_ELTCALCMSA' + IntToStr(NumElement), '12');
            TC.PutValue('PE3_MONTANTELT' + IntToStr(NumElement), Arrondi(Montant12, 0));
          end;
          //FIN PT13
{PT18  On ne le fait plus car on ne sait pas détecter l'arret de travail
          if Montant13 <> 0 then
          begin
            NumElement := NumElement + 1;
            TC.PutValue('PE3_ELTCALCMSA' + IntToStr(NumElement), '13');
            TC.PutValue('PE3_MONTANTELT' + IntToStr(NumElement), Arrondi(Montant13, 0));
          end;
}
          if Montant14 <> 0 then
          begin
            NumElement := NumElement + 1;
            TC.PutValue('PE3_ELTCALCMSA' + IntToStr(NumElement), '14');
            TC.PutValue('PE3_MONTANTELT' + IntToStr(NumElement), Arrondi(Montant14, 0));
          end;
          if NumElement = 5 then
          begin
            TC.InsertorUpdateDB(False);
            TC.PutValue('PE3_ORDRE', 1);
            NumElement := 0;
          end;
          if Montant81 <> 0 then
          begin
            NumElement := NumElement + 1;
            TC.PutValue('PE3_ELTCALCMSA' + IntToStr(NumElement), '81');
            TC.PutValue('PE3_MONTANTELT' + IntToStr(NumElement), Arrondi(Montant81, 0));
          end;
          if NumElement = 5 then
          begin
            TC.InsertorUpdateDB(False);
            TC.PutValue('PE3_ORDRE', 1);
            NumElement := 0;
          end;
          if Montant82 <> 0 then
          begin
            NumElement := NumElement + 1;
            TC.PutValue('PE3_ELTCALCMSA' + IntToStr(NumElement), '82');
            TC.PutValue('PE3_MONTANTELT' + IntToStr(NumElement), Arrondi(Montant82, 0));
          end;
          if NumElement = 5 then
          begin
            TC.InsertorUpdateDB(False);
            TC.PutValue('PE3_ORDRE', 1);
            NumElement := 0;
          end;
          if Montant83 <> 0 then
          begin
            NumElement := NumElement + 1;
            TC.PutValue('PE3_ELTCALCMSA' + IntToStr(NumElement), '83');
            TC.PutValue('PE3_MONTANTELT' + IntToStr(NumElement), Arrondi(Montant83, 0));
          end;
          for x := NumElement + 1 to 5 do
          begin
            TC.PutValue('PE3_ELTCALCMSA' + IntToStr(x), '');
            TC.PutValue('PE3_MONTANTELT' + IntToStr(x), 0);
          end;
//PT22          if NumElement = 0 then Writeln(FRapport, '- Aucune rémunération pour la période du ' + DateToStr(debPaie) + ' au ' + DateToStr(FinPaie) + ' est égale à 0');
          TC.InsertorUpdateDB(False);

          //DEB PT14
          //DEB PT16
          //Supprimer les enregistrements déjà présents
{PT19 Un seul delete en début          ExecuteSQL('DELETE FROM MSAPERIODESPE31' +
            ' WHERE PE3_SALARIE="' + Salarie + '"' +
            ' AND PE3_DATEDEBUT="' + USDATETIME(TobPaie.Detail[i].GetValue('PPU_DATEDEBUT')) + '"' +
            ' AND PE3_DATEFIN="' + USDATETIME(TobPaie.Detail[i].GetValue('PPU_DATEFIN')) + '"' +
            ' AND PE3_ORDRE >= 100');
}
          //FIN PT16
          TC := Tob.Create('MSAPERIODESPE31', TobMSAPeriodes, -1);
          TC.PutValue('PE3_SALARIE', Salarie);
          TC.PutValue('PE3_NOM', Nom);
          TC.PutValue('PE3_PRENOM', Prenom);
          TC.PutValue('PE3_DATENAISSANCE', DateNaiss);
          TC.PutValue('PE3_MSAACTIVITE', ActiviteMsa);
          TC.PutValue('PE3_NUMEROSS', NumSS);
          TC.PutValue('PE3_ETABLISSEMENT', TobPaie.Detail[i].GetValue('PPU_ETABLISSEMENT'));
          TC.PutValue('PE3_DATEDEBUT', TobPaie.Detail[i].GetValue('PPU_DATEDEBUT'));
          TC.PutValue('PE3_DATEFIN', TobPaie.Detail[i].GetValue('PPU_DATEFIN'));
          TC.PutValue('PE3_ORDRE', 100);
          TC.PutValue('PE3_UGMSA', UGMsa);
          TC.PutValue('PE3_TOPPLAF', '-');
          TC.PutValue('PE3_PEXOMSA', '');
          TC.PutValue('PE3_NBJOURS', 0);
          TC.PutValue('PE3_NBHEURES', 0);

          For r := 1 to 5 do
            TC.PutValue('PE3_ELTCALCMSA' + IntToStr(r), '');

          NbHeuresSup := 0;
          for r := 0 to 3 do
          begin
            if r = 0 then Cumul := '24';
            if r = 1 then Cumul := '26';
            if r = 2 then Cumul := GetParamSocSecur('SO_PGMSACUMUL01','');
            if r = 3 then Cumul := GetParamSocSecur('SO_PGMSACUMUL02','');

            if Cumul <> '' then
            begin
              Q := OpenSQL('SELECT SUM(PHB_BASEREM) AS PHB_BASEREM,SUM(PHB_COEFFREM) AS PHB_COEFFREM,' +  //PT16 cumuler les valeurs
                ' SUM(PHB_MTREM) AS PHB_MTREM ,COUNT(*) AS COMPTEUR FROM HISTOBULLETIN ' +                //PT16 cumuler les valeurs
                ' WHERE PHB_SALARIE = "' + Salarie + '" AND ' +
                ' PHB_DATEFIN>="' + UsDateTime(DebPaie) + '" AND' +
                ' PHB_DATEFIN<="' + UsDateTime(FinPaie) + '"' +
                ' AND PHB_RUBRIQUE IN (SELECT DISTINCT PCR_RUBRIQUE FROM CUMULRUBRIQUE ' +
                ' WHERE PCR_CUMULPAIE = "' + Cumul + '") AND PHB_COEFFREM <> 0', True);
              if (not Q.Eof) and (Q.FindField('PHB_BASEREM').AsFloat <> 0) then
              begin
                TC.PutValue('PE3_ORDRE', 100 + r);
//PT18                if ((NbHeuresSup + Q.FindField('PHB_BASEREM').AsFloat) <= 8) then //PT16
                if (NbHeuresSup = 0) then //PT16  //PT18
                  TC.PutValue('PE3_MONTANTELT1', '30') // Type rémunération
                else
                  if (Cumul = '44') then    //PT20
                    TC.PutValue('PE3_MONTANTELT1', '30')  //PT20
                  else
                    TC.PutValue('PE3_MONTANTELT1', '40'); // Type rémunération
                NbHeuresSup := NbHeuresSup + Q.FindField('PHB_BASEREM').AsFloat;
                TC.PutValue('PE3_MONTANTELT2', ((Q.FindField('PHB_COEFFREM').AsFloat / Q.FindField('COMPTEUR').AsInteger) * 100) - 100); // Taux majoration  //PT16
                TC.PutValue('PE3_MONTANTELT3', Q.FindField('PHB_BASEREM').AsFloat * 100); // Nombre d'heures
                TC.PutValue('PE3_MONTANTELT4', Q.FindField('PHB_MTREM').AsFloat); // Montant rémunération
                TC.PutValue('PE3_MONTANTELT5', 0);
                TC.InsertorUpdateDB(False);
              end;
              Ferme(Q);
            end;
          end;
          //FIN PT14

          //DEB PT22
          if NbHeuresSup <> 0 then
          begin
            Cumul := GetParamSocSecur('SO_PGMSACUMULCOT','');
            Q := OpenSQL('SELECT SUM(PHB_MTSALARIAL) AS PHB_MTSALARIAL FROM HISTOBULLETIN ' +
              ' WHERE PHB_SALARIE = "' + Salarie + '" AND ' +
              ' PHB_DATEFIN>="' + UsDateTime(DebPaie) + '" AND' +
              ' PHB_DATEFIN<="' + UsDateTime(FinPaie) + '"' +
              ' AND PHB_RUBRIQUE IN (SELECT DISTINCT PCR_RUBRIQUE FROM CUMULRUBRIQUE ' +
              ' WHERE PCR_CUMULPAIE = "' + Cumul + '")' , True);  {PT22 AND PHB_COEFFREM <> 0'}
            if (not Q.Eof) and (Q.FindField('PHB_MTSALARIAL').AsFloat <> 0) then
            begin
              TC := Tob.Create('MSAPERIODESPE31', TobMSAPeriodes, -1);
              TC.PutValue('PE3_SALARIE', Salarie);
              TC.PutValue('PE3_NOM', Nom);
              TC.PutValue('PE3_PRENOM', Prenom);
              TC.PutValue('PE3_DATENAISSANCE', DateNaiss);
              TC.PutValue('PE3_MSAACTIVITE', ActiviteMsa);
              TC.PutValue('PE3_NUMEROSS', NumSS);
              TC.PutValue('PE3_ETABLISSEMENT', TobPaie.Detail[i].GetValue('PPU_ETABLISSEMENT'));
              TC.PutValue('PE3_DATEDEBUT', TobPaie.Detail[i].GetValue('PPU_DATEDEBUT'));
              TC.PutValue('PE3_DATEFIN', TobPaie.Detail[i].GetValue('PPU_DATEFIN'));
              TC.PutValue('PE3_ORDRE', 0);
              TC.PutValue('PE3_UGMSA', UGMsa);
              TC.PutValue('PE3_TOPPLAF', '-');
              TC.PutValue('PE3_PEXOMSA', '');  //PT15
              TC.PutValue('PE3_NBJOURS', Jours);
              TC.PutValue('PE3_NBHEURES', NbHeures);
              if NumElement = 5 then
              begin
                TC.InsertorUpdateDB(False);
                TC.PutValue('PE3_ORDRE', 1);
                NumElement := 0;
              end;
              NumElement := NumElement + 1;
              TC.PutValue('PE3_ELTCALCMSA' + IntToStr(NumElement), '25');
              TC.PutValue('PE3_MONTANTELT' + IntToStr(NumElement), Q.FindField('PHB_MTSALARIAL').AsFloat);
              for x := NumElement + 1 to 5 do
              begin
                TC.PutValue('PE3_ELTCALCMSA' + IntToStr(x), '');
                TC.PutValue('PE3_MONTANTELT' + IntToStr(x), 0);
              end;
              TC.InsertorUpdateDB(False);
            end;
            Ferme(Q);
          end;
          if NumElement = 0 then Writeln(FRapport, '- Aucune rémunération pour la période du ' + DateToStr(debPaie) + ' au ' + DateToStr(FinPaie) + ' est égale à 0');
          //FIN PT22

          if MajEvolutions = True then //Si recherche évolutions du salarié
          begin
            //Enregistrement PE21 Dates entrée et sortie
            DateEntree := TS.GetValue('PSA_DATEENTREE');
            DateSortie := TS.GetValue('PSA_DATESORTIE');
            if (DateEntree >= DateDebut) and (DateEntree <= DateFin) then
            begin
              St := 'SELECT PE2_SALARIE FROM MSAEVOLUTIONSPE2 WHERE ' +
                'PE2_TYPEEVOLMSA="PE21" AND PE2_SALARIE="' + Salarie + '" AND PE2_DEBACTIVITE="' + UsDateTime(DateEntree) + '"';
              TobEvolutions := Tob.Create('MSAEVOLUTIONSPE2', nil, -1);
              TobEvolutions.LoadDetailDBFromSQL ('MSAEVOLUTIONSPE2', St);
              TE := TobEvolutions.FindFirst(['PE2_SALARIE'], [Salarie], False);
              if TE = nil then
              begin
                TE := Tob.Create('MSAEVOLUTIONSPE2', TobEvolutions, -1);
                TE.PutValue('PE2_ETABLISSEMENT', EtablissementS);
                TE.PutValue('PE2_SALARIE', Salarie);
                TE.PutValue('PE2_NOM', Nom);
                TE.PutValue('PE2_PRENOM', Prenom);
                TE.PutValue('PE2_DATENAISSANCE', DateNaiss);
                TE.PutValue('PE2_MSAACTIVITE', ActiviteMsa);  //PT10
                TE.PutValue('PE2_NUMEROSS', NumSS);
                TE.PutValue('PE2_TYPEEVOLMSA', 'PE21');
                TE.PutValue('PE2_DATEEFFET', DateEntree);
                TE.PutValue('PE2_DEBACTIVITE', DateEntree);
                If DateEntree = DateSortie then TE.PutValue('PE2_FINACTIVITE', DateSortie); //PT8
                TE.InsertDB(nil, False);
              end;
              FreeAndNil(TobEvolutions);
            end;
            if (DateSortie >= DateDebut) and (DateSortie <= DateFin) and (DateSortie <> DateEntree) then //PT8
            begin
              St := 'SELECT PE2_SALARIE FROM MSAEVOLUTIONSPE2 WHERE ' +
                'PE2_TYPEEVOLMSA="PE21" AND PE2_SALARIE="' + Salarie + '" AND PE2_FINACTIVITE="' + UsDateTime(DateSortie) + '"';
              TobEvolutions := Tob.Create('MSAEVOLUTIONSPE2', nil, -1);
              TobEvolutions.LoadDetailDBFromSQL ('MSAEVOLUTIONSPE2', St);
              TE := TobEvolutions.FindFirst(['PE2_SALARIE'], [Salarie], False);
              if TE = nil then
              begin
                TE := Tob.Create('MSAEVOLUTIONSPE2', TobEvolutions, -1);
                TE.PutValue('PE2_ETABLISSEMENT', EtablissementS);
                TE.PutValue('PE2_SALARIE', Salarie);
                TE.PutValue('PE2_NOM', Nom);
                TE.PutValue('PE2_PRENOM', Prenom);
                TE.PutValue('PE2_DATENAISSANCE', DateNaiss);
                TE.PutValue('PE2_MSAACTIVITE', ActiviteMsa); //PT10
                TE.PutValue('PE2_NUMEROSS', NumSS);
                TE.PutValue('PE2_TYPEEVOLMSA', 'PE21');
                TE.PutValue('PE2_DATEEFFET', DateSortie);
                TE.PutValue('PE2_FINACTIVITE', DateSortie);
                TE.InsertDB(nil, False);
              end;
              FreeAndNil(TobEvolutions);
            end;
          end;

          Writeln(FRapport, '');
        end;

        //DEBUT PT12
        St := 'SELECT * FROM ABSENCESALARIE'+
         ' LEFT JOIN MOTIFABSENCE ON ##PMA_PREDEFINI##'+
        ' PCN_TYPECONGE=PMA_MOTIFABSENCE '+
        'WHERE PCN_SALARIE="'+Salarie+'" AND (PMA_TYPEABS="MAN" OR PMA_TYPEABS="MAP") AND '+
        'PCN_DATEDEBUTABS>="'+UsDateTime(DateDebut)+'" AND PCN_DATEDEBUTABS<="'+UsDateTime(DateFin)+'"';
        TobAbs := Tob.Create('les absences',Nil,-1);
        TobAbs.LoadDetailDBFromSQL ('les absences', St);
        For NbAbs := 0 to TobAbs.Detail.Count - 1 do
        begin
          St := 'SELECT PE2_SALARIE FROM MSAEVOLUTIONSPE2 WHERE ' +
            'PE2_TYPEEVOLMSA="PE22" AND PE2_DATEEFFET="'+UsDateTime(TobAbs.Detail[NbAbs].GetValue('PCN_DATEDEBUTABS'))+'" AND '+
            'PE2_SALARIE="' + Salarie + '" AND PE2_FINACTIVITE="' + UsDateTime(TS.GetValue('PSA_DATESORTIE')) + '"';
          TobEvolutions := Tob.Create('MSAEVOLUTIONSPE2', nil, -1);
          TobEvolutions.LoadDetailDBFromSQL ('MSAEVOLUTIONSPE2', St);
          TE := TobEvolutions.FindFirst(['PE2_SALARIE'], [Salarie], False);
          if TE = nil then TE := Tob.Create('MSAEVOLUTIONSPE2', TobEvolutions, -1);
          TE.PutValue('PE2_ETABLISSEMENT', EtablissementS);
          TE.PutValue('PE2_SALARIE', Salarie);
          TE.PutValue('PE2_NOM', Nom);
          TE.PutValue('PE2_PRENOM', Prenom);
          TE.PutValue('PE2_DATENAISSANCE', DateNaiss);
          TE.PutValue('PE2_MSAACTIVITE', ActiviteMsa); //PT10
          TE.PutValue('PE2_NUMEROSS', NumSS);
          TE.PutValue('PE2_TYPEEVOLMSA', 'PE22');
          TE.PutValue('PE2_DATEEFFET', TobAbs.Detail[NbAbs].GetValue('PCN_DATEDEBUTABS'));
          TE.PutValue('PE2_DEBSUSPCT', TobAbs.Detail[NbAbs].GetValue('PCN_DATEDEBUTABS'));
          TE.PutValue('PE2_FINSUSSPCT', TobAbs.Detail[NbAbs].GetValue('PCN_DATEFINABS'));
          TE.PutValue('PE2_MSASUSPCT', '06');
          TE.InsertOrUpdateDB(False);
          FreeAndNil(TobEvolutions);
        end;
        FreeAndNil(TobAbs);
        //FIN PT12
        FreeAndNil(TobPaie);
        FreeAndNil(TobMSAPeriodes);
        MoveCurProgressForm(Nom);
      end;
      FiniMoveProgressForm;
    end;
    if liste.AllSelected then
    begin
      InitMoveProgressForm(nil, 'Début du traitement', 'Veuillez patienter SVP ...', TFmul(Ecran).Q.RecordCount, FALSE, TRUE);
      InitMove(TFmul(Ecran).Q.RecordCount, '');
      Q_Mul.First;
      while not Q_Mul.EOF do
      begin
        Salarie := TFmul(Ecran).Q.FindField('PSE_SALARIE').asstring;
        TS := TobSalaries.FindFirst(['PSA_SALARIE'], [Salarie], False);
        EtablissementS := TS.GetValue('PSA_ETABLISSEMENT');
        Nom := TS.GetValue('PSA_LIBELLE');
        Prenom := TS.GetValue('PSA_PRENOM');
        Writeln(FRapport, '');
        Writeln(FRapport, '     Salarie :' + Salarie + ' ' + Nom + ' ' + Prenom);
        DateNaiss := TS.GetValue('PSA_DATENAISSANCE');
        NumSS := TS.GetValue('PSA_NUMEROSS');
        Sexe := TS.GetValue('PSA_SEXE');
        ActiviteMsa := TS.Getvalue('PSE_MSAACTIVITE');
        UGMsa := TS.Getvalue('PSE_MSAUNITEGES');
        if TestNumeroSS(NumSS, Sexe) < 0 then Writeln(FRapport, '- véifier le N° de sécurité sociale');
        if DateNaiss = IDate1900 then Writeln(FRapport, '- la date de naissance n''est pas renseignée');
        if ActiviteMsa = '' then Writeln(FRapport, '- l''activité n''est pas renseignée');
        St := 'SELECT PPU_SALARIE,PPU_ETABLISSEMENT,PPU_DATEDEBUT,PPU_DATEFIN,PPU_PAYELE,PPU_CBRUT,(PHC_MONTANT) AS HEURESTRAV' +
          ' FROM PAIEENCOURS' +
          ' LEFT JOIN HISTOCUMSAL ON PPU_DATEDEBUT=PHC_DATEDEBUT AND PPU_DATEFIN=PHC_DATEFIN AND PPU_SALARIE=PHC_SALARIE AND PHC_CUMULPAIE="21"' +
          ' WHERE PPU_SALARIE="' + Salarie + '" AND' +
          ' PPU_DATEDEBUT>="' + UsDatetime(DateDebut) + '"' +
          ' AND PPU_DATEFIN<="' + UsDatetime(DateFin) + '"';
        TobPaie := Tob.Create('Les paies', nil, -1);
        TobPaie.LoadDetailDBFromSQL ('PAIEENCOURS', St);

        //DEB PT19
        //Supprimer les enregistrements déjà présents car la remise à 0 d'un champ par une tob
        //ne fonctionne pas
        St := 'DELETE FROM MSAPERIODESPE31' +
          ' WHERE PE3_SALARIE="' + Salarie + '"' +
          ' AND PE3_DATEDEBUT>="' + USDATETIME(DateDebut) + '"' +
          ' AND PE3_DATEFIN<="' + USDATETIME(DateFin) + '"';
        ExecuteSQL(St);
        //FIN PT19

        St := 'SELECT * FROM MSAPERIODESPE31 WHERE PE3_SALARIE="' + Salarie + '"' +
          ' AND PE3_DATEDEBUT>="' + UsDatetime(DateDebut) + '"' +
          ' AND PE3_DATEFIN<="' + UsDatetime(DateFin) + '"';
        TobMSAPeriodes := Tob.create('MSAPERIODESPE31', nil, -1);
        TobMSAPeriodes.LoadDetailDBFromSQL ('MSAPERIODESPE31', St);
        if TobPaie.Detail.Count <= 0 then Writeln(FRapport, '- Ce salarié n''a aucune période de paie');
        for i := 0 to TobPaie.Detail.Count - 1 do
        begin
          DebPaie := TobPaie.Detail[i].GetValue('PPU_DATEDEBUT');
          FinPaie := TobPaie.Detail[i].GetValue('PPU_DATEFIN');
          If TobPaie.Detail[i].GetValue('HEURESTRAV') <> null then
          begin
               if IsNumeric(TobPaie.Detail[i].GetValue('HEURESTRAV')) then NbHeures := TobPaie.Detail[i].GetValue('HEURESTRAV')  // PT1 et PT3
               else NbHeures := 0;
          end
          else NbHeures := 0;
          if NbHeures = 0 then Writeln(FRapport, '- Le nombre d''heures de la période du ' + DateToStr(debPaie) + ' au ' + DateToStr(FinPaie) + ' est égale à 0');  //PT23
          //                                        Montant01 := TobPaie.Detail[i].GetValue('PPU_CBRUT');
          Montant01 := 0;
          Montant02 := 0;
          Montant09 := 0;
          Montant12 := 0;
          Montant13 := 0;
          Montant14 := 0;
          Montant81 := 0;
          Montant82 := 0;
          Montant83 := 0;
          MontantR5850 := 0; //PT17
          //DEBUT PT5
          Q := OpenSQL('SELECT SUM('+ChampBul+') MONTANT FROM HISTOBULLETIN WHERE PHB_SALARIE="'+Salarie+'" '+
          'AND PHB_NATURERUB="AAA" AND PHB_RUBRIQUE="'+RubJours+'" '+
          'AND PHB_DATEDEBUT>="'+UsDateTime(DebPaie)+'" '+
          'AND PHB_DATEFIN<="'+UsDateTime(FinPaie)+'"',True);
          If Not Q.Eof then Jours := Q.FindField('MONTANT').AsFloat
          else Jours := 0;
          Ferme(Q);
          //FIN PT5
          TC := Tob.Create('MSAPERIODESPE31', TobMSAPeriodes, -1);
          TC.PutValue('PE3_SALARIE', Salarie);
          TC.PutValue('PE3_NOM', Nom);
          TC.PutValue('PE3_PRENOM', Prenom);
          TC.PutValue('PE3_DATENAISSANCE', DateNaiss);
          TC.PutValue('PE3_MSAACTIVITE', ActiviteMsa);
          TC.PutValue('PE3_NUMEROSS', NumSS);
          TC.PutValue('PE3_ETABLISSEMENT', TobPaie.Detail[i].GetValue('PPU_ETABLISSEMENT'));
          TC.PutValue('PE3_DATEDEBUT', TobPaie.Detail[i].GetValue('PPU_DATEDEBUT'));
          TC.PutValue('PE3_DATEFIN', TobPaie.Detail[i].GetValue('PPU_DATEFIN'));
          TC.PutValue('PE3_ORDRE', 0);
          TC.PutValue('PE3_UGMSA', UGMsa);
          TC.PutValue('PE3_TOPPLAF', '-');
          TC.PutValue('PE3_PEXOMSA', '');
          TC.PutValue('PE3_NBJOURS', Jours);
          TC.PutValue('PE3_NBHEURES', NbHeures);
          For r := 1 to 5 do
          begin
            TC.PutValue('PE3_ELTCALCMSA' + IntToStr(r), '');
            TC.PutValue('PE3_MONTANTELT' + IntToStr(r), 0);
          end;
          Q := OpenSQL('SELECT (PHC_MONTANT) MONTANT' +
            ' FROM HISTOCUMSAL' +
            ' WHERE PHC_SALARIE="' + Salarie + '" AND' +
            ' PHC_CUMULPAIE="02" AND' +
            ' PHC_DATEDEBUT>="' + UsDatetime(DebPaie) + '"' +
            ' AND PHC_DATEFIN<="' + UsDatetime(FinPaie) + '"', True);
          if not Q.Eof then Montant01 := Q.FindField('MONTANT').AsFloat;
          Ferme(Q);
          Q := OpenSQL('SELECT SUM(PHB_MTREM) AS MTPREAVIS' +
            ' FROM HISTOBULLETIN' +
            ' LEFT JOIN REMUNERATION ON' +
            ' PHB_NATURERUB=PRM_NATURERUB AND' +
            ' ##PRM_PREDEFINI## PHB_RUBRIQUE=PRM_RUBRIQUE WHERE' +
            ' PRM_INDEMPREAVIS="X" AND' +
            ' PHB_SALARIE="' + Salarie + '" AND' +
            ' PHB_DATEFIN>="' + UsDateTime(DebPaie) + '" AND' +
            ' PHB_DATEFIN<="' + UsDateTime(FinPaie) + '"', True);
          if not Q.Eof then Montant02 := Q.FindField('MTPREAVIS').AsFloat;
          Ferme(Q);
          Q := OpenSQL('SELECT SUM(PHB_MTREM) AS MTCP' +
            ' FROM HISTOBULLETIN' +
            ' LEFT JOIN REMUNERATION ON' +
            ' PHB_NATURERUB=PRM_NATURERUB AND' +
            ' ##PRM_PREDEFINI## PHB_RUBRIQUE=PRM_RUBRIQUE WHERE' +
            ' PRM_INDEMCOMPCP="X" AND' +
            ' PHB_SALARIE="' + Salarie + '" AND' +
            ' PHB_DATEFIN>="' + UsDateTime(DebPaie) + '" AND' +
            ' PHB_DATEFIN<="' + UsDateTime(FinPaie) + '"', True);
          if not Q.Eof then Montant09 := Q.FindField('MTCP').AsFloat;
          Ferme(Q);
          //DEBUT PT13
          Q := OpenSQL('SELECT SUM(PHB_MTREM) AS MTCP' +
            ' FROM HISTOBULLETIN' +
            ' LEFT JOIN REMUNERATION ON' +
            ' PHB_NATURERUB=PRM_NATURERUB AND' +
            ' ##PRM_PREDEFINI## PHB_RUBRIQUE=PRM_RUBRIQUE WHERE' +
            ' PRM_RUBRIQUE="6000" AND' +
            ' PHB_SALARIE="' + Salarie + '" AND' +
            ' PHB_DATEFIN>="' + UsDateTime(DebPaie) + '" AND' +
            ' PHB_DATEFIN<="' + UsDateTime(FinPaie) + '"', True);
          if not Q.Eof then Montant12 := Q.FindField('MTCP').AsFloat;
          Ferme(Q);
          //FIN PT13
          Q := OpenSQL('SELECT (PHC_MONTANT) MONTANT' +
            ' FROM HISTOCUMSAL' +
            ' WHERE PHC_SALARIE="' + Salarie + '" AND' +
            ' PHC_CUMULPAIE="05" AND' +
            ' PHC_DATEDEBUT>="' + UsDatetime(DebPaie) + '"' +
            ' AND PHC_DATEFIN<="' + UsDatetime(FinPaie) + '"', True);
          if not Q.Eof then Montant13 := Q.FindField('MONTANT').AsFloat;
          Ferme(Q);

{PT18          //DEB PT17
          Q := OpenSQL('SELECT PHB_MTPATRONAL FROM HISTOBULLETIN' +
            ' WHERE PHB_SALARIE="' + Salarie + '" AND ' +
            ' PHB_NATURERUB="COT" AND PHB_RUBRIQUE="5850" AND ' +
            ' PHB_DATEDEBUT>="' + UsDatetime(DebPaie) + '"' +
            ' AND PHB_DATEFIN<="' + UsDatetime(FinPaie) + '"', True);
          if not Q.Eof then MontantR5850 := Q.FindField('PHB_MTPATRONAL').AsFloat;
          Ferme(Q);
          //FIN PT17
}
//PT18          if (Montant13 <> 0) and ((Montant13 = Montant01) or (MontantR5850 >= 0)) then Montant13 := 0;  //PT17
          if (Montant13 = Montant01) then Montant13 := 0;  //PT18

          Q := OpenSQL('SELECT (PHC_MONTANT) MONTANT' +
            ' FROM HISTOCUMSAL' +
            ' WHERE PHC_SALARIE="' + Salarie + '" AND' +
            ' PHC_CUMULPAIE="39" AND' +
            ' PHC_DATEDEBUT>="' + UsDatetime(DebPaie) + '"' +
            ' AND PHC_DATEFIN<="' + UsDatetime(FinPaie) + '"', True);
          if not Q.Eof then Montant14 := Q.FindField('MONTANT').AsFloat;
          Ferme(Q);
          if TS.GetValue('PSA_CONDEMPLOI') = 'P' then
          begin
            Montant81 := TS.GetValue('PSA_SALAIRETHEO');
//PT21            Montant82 := Montant13;
//PT18            Montant13 := 0;
          end;
          Montant13 := 0;  //PT18
          if TS.GetValue('PSA_MULTIEMPLOY') = 'X' then
          begin
            Montant83 := TS.GetValue('PSA_SALAIREMULTI');
          end;
          Montant01 := Montant01 - Montant02 - Montant09 - Montant14;
          NumElement := 0;
//          if Montant01 <> 0 then PT4
 //         begin
            NumElement := NumElement + 1;
            TC.PutValue('PE3_ELTCALCMSA' + IntToStr(NumElement), '01');
            TC.PutValue('PE3_MONTANTELT' + IntToStr(NumElement), Arrondi(Montant01, 0));
 //         end;
          if Montant02 <> 0 then
          begin
            NumElement := NumElement + 1;
            TC.PutValue('PE3_ELTCALCMSA' + IntToStr(NumElement), '02');
            TC.PutValue('PE3_MONTANTELT' + IntToStr(NumElement), Arrondi(Montant02, 0));
          end;
          if Montant09 <> 0 then
          begin
            NumElement := NumElement + 1;
            TC.PutValue('PE3_ELTCALCMSA' + IntToStr(NumElement), '09');
            TC.PutValue('PE3_MONTANTELT' + IntToStr(NumElement), Arrondi(Montant09, 0));
          end;
          //DEBUT PT13
          if Montant12 <> 0 then
          begin
            NumElement := NumElement + 1;
            TC.PutValue('PE3_ELTCALCMSA' + IntToStr(NumElement), '12');
            TC.PutValue('PE3_MONTANTELT' + IntToStr(NumElement), Arrondi(Montant12, 0));
          end;
          //FIN PT13
{PT18 On ne le fait pas car on ne sait pas détecter l'arret de travail
          if Montant13 <> 0 then
          begin
            NumElement := NumElement + 1;
            TC.PutValue('PE3_ELTCALCMSA' + IntToStr(NumElement), '13');
            TC.PutValue('PE3_MONTANTELT' + IntToStr(NumElement), Arrondi(Montant13, 0));
          end;
}
          if Montant14 <> 0 then
          begin
            NumElement := NumElement + 1;
            TC.PutValue('PE3_ELTCALCMSA' + IntToStr(NumElement), '14');
            TC.PutValue('PE3_MONTANTELT' + IntToStr(NumElement), Arrondi(Montant14, 0));
          end;
          if NumElement = 5 then
          begin
            TC.InsertorUpdateDB(False);
            TC.PutValue('PE3_ORDRE', 1);
            NumElement := 0;
          end;
          if NumElement = 5 then
          begin
            TC.InsertorUpdateDB(False);
            TC.PutValue('PE3_ORDRE', 1);
            NumElement := 0;
          end;
          if Montant81 <> 0 then
          begin
            NumElement := NumElement + 1;
            TC.PutValue('PE3_ELTCALCMSA' + IntToStr(NumElement), '81');
            TC.PutValue('PE3_MONTANTELT' + IntToStr(NumElement), Arrondi(Montant81, 0));
          end;
          if NumElement = 5 then
          begin
            TC.InsertorUpdateDB(False);
            TC.PutValue('PE3_ORDRE', 1);
            NumElement := 0;
          end;
          if Montant82 <> 0 then
          begin
            NumElement := NumElement + 1;
            TC.PutValue('PE3_ELTCALCMSA' + IntToStr(NumElement), '82');
            TC.PutValue('PE3_MONTANTELT' + IntToStr(NumElement), Arrondi(Montant82, 0));
          end;
          if NumElement = 5 then
          begin
            TC.InsertorUpdateDB(False);
            TC.PutValue('PE3_ORDRE', 1);
            NumElement := 0;
          end;
          if Montant83 <> 0 then
          begin
            NumElement := NumElement + 1;
            TC.PutValue('PE3_ELTCALCMSA' + IntToStr(NumElement), '83');
            TC.PutValue('PE3_MONTANTELT' + IntToStr(NumElement), Arrondi(Montant83, 0));
          end;
          for x := NumElement + 1 to 5 do
          begin
            TC.PutValue('PE3_ELTCALCMSA' + IntToStr(x), '');
            TC.PutValue('PE3_MONTANTELT' + IntToStr(x), 0);
          end;
//PT22          if NumElement = 0 then Writeln(FRapport, '- Aucune rémunération pour la période du ' + DateToStr(debPaie) + ' au ' + DateToStr(FinPaie) + ' est égale à 0');
          TC.InsertorUpdateDB(False);

          //DEB PT14
          //DEB PT16
{PT19          //Supprimer les enregistrements déjà présents
          ExecuteSQL('DELETE FROM MSAPERIODESPE31' +
            ' WHERE PE3_SALARIE="' + Salarie + '"' +
            ' AND PE3_DATEDEBUT="' + USDATETIME(TobPaie.Detail[i].GetValue('PPU_DATEDEBUT')) + '"' +
            ' AND PE3_DATEFIN="' + USDATETIME(TobPaie.Detail[i].GetValue('PPU_DATEFIN')) + '"' +
            ' AND PE3_ORDRE >= 100');
}
          //FIN PT16
          TC := Tob.Create('MSAPERIODESPE31', TobMSAPeriodes, -1);
          TC.PutValue('PE3_SALARIE', Salarie);
          TC.PutValue('PE3_NOM', Nom);
          TC.PutValue('PE3_PRENOM', Prenom);
          TC.PutValue('PE3_DATENAISSANCE', DateNaiss);
          TC.PutValue('PE3_MSAACTIVITE', ActiviteMsa);
          TC.PutValue('PE3_NUMEROSS', NumSS);
          TC.PutValue('PE3_ETABLISSEMENT', TobPaie.Detail[i].GetValue('PPU_ETABLISSEMENT'));
          TC.PutValue('PE3_DATEDEBUT', TobPaie.Detail[i].GetValue('PPU_DATEDEBUT'));
          TC.PutValue('PE3_DATEFIN', TobPaie.Detail[i].GetValue('PPU_DATEFIN'));
          TC.PutValue('PE3_ORDRE', 100);
          TC.PutValue('PE3_UGMSA', UGMsa);
          TC.PutValue('PE3_TOPPLAF', '-');
          TC.PutValue('PE3_PEXOMSA', '');
          TC.PutValue('PE3_NBJOURS', 0);
          TC.PutValue('PE3_NBHEURES', 0);

          For r := 1 to 5 do
            TC.PutValue('PE3_ELTCALCMSA' + IntToStr(r), '');

          NbHeuresSup := 0;
          for r := 0 to 3 do
          begin
            if r = 0 then Cumul := '24';
            if r = 1 then Cumul := '26';
            if r = 2 then Cumul := GetParamSocSecur('SO_PGMSACUMUL01','');
            if r = 3 then Cumul := GetParamSocSecur('SO_PGMSACUMUL02','');

            if Cumul <> '' then
            begin
              Q := OpenSQL('SELECT SUM(PHB_BASEREM) AS PHB_BASEREM,SUM(PHB_COEFFREM) AS PHB_COEFFREM,' +  //PT16 cumuler les valeurs
                ' SUM(PHB_MTREM) AS PHB_MTREM ,COUNT(*) AS COMPTEUR FROM HISTOBULLETIN ' +                //PT16 cumuler les valeurs
                ' WHERE PHB_SALARIE = "' + Salarie + '" AND ' +
                ' PHB_DATEFIN>="' + UsDateTime(DebPaie) + '" AND' +
                ' PHB_DATEFIN<="' + UsDateTime(FinPaie) + '"' +
                ' AND PHB_RUBRIQUE IN (SELECT DISTINCT PCR_RUBRIQUE FROM CUMULRUBRIQUE ' +
                ' WHERE PCR_CUMULPAIE = "' + Cumul + '") AND PHB_COEFFREM <> 0', True);
              if (not Q.Eof) and (Q.FindField('PHB_BASEREM').AsFloat <> 0) then
              begin
                TC.PutValue('PE3_ORDRE', 100 + r);
                if (NbHeuresSup = 0) then //PT16  //PT18
                  TC.PutValue('PE3_MONTANTELT1', '30') // Type rémunération
                else
                  if (Cumul = '44') then    //PT20
                    TC.PutValue('PE3_MONTANTELT1', '30')  //PT20
                  else
                    TC.PutValue('PE3_MONTANTELT1', '40'); // Type rémunération
                NbHeuresSup := NbHeuresSup + Q.FindField('PHB_BASEREM').AsFloat;
                TC.PutValue('PE3_MONTANTELT2', ((Q.FindField('PHB_COEFFREM').AsFloat / Q.FindField('COMPTEUR').AsInteger) * 100) - 100); // Taux majoration  //PT16
                TC.PutValue('PE3_MONTANTELT3', Q.FindField('PHB_BASEREM').AsFloat * 100); // Nombre d'heures
                TC.PutValue('PE3_MONTANTELT4', Q.FindField('PHB_MTREM').AsFloat); // Montant rémunération
                TC.PutValue('PE3_MONTANTELT5', 0);
                TC.InsertorUpdateDB(False);
              end;
              Ferme(Q);
            end;
          end;
          //FIN PT14

          //DEB PT22
          if NbHeuresSup <> 0 then
          begin
            Cumul := GetParamSocSecur('SO_PGMSACUMULCOT','');
            Q := OpenSQL('SELECT SUM(PHB_MTSALARIAL) AS PHB_MTSALARIAL FROM HISTOBULLETIN ' +
              ' WHERE PHB_SALARIE = "' + Salarie + '" AND ' +
              ' PHB_DATEFIN>="' + UsDateTime(DebPaie) + '" AND' +
              ' PHB_DATEFIN<="' + UsDateTime(FinPaie) + '"' +
              ' AND PHB_RUBRIQUE IN (SELECT DISTINCT PCR_RUBRIQUE FROM CUMULRUBRIQUE ' +
              ' WHERE PCR_CUMULPAIE = "' + Cumul + '")', True);
            if (not Q.Eof) and (Q.FindField('PHB_MTSALARIAL').AsFloat <> 0) then
            begin
              TC := Tob.Create('MSAPERIODESPE31', TobMSAPeriodes, -1);
              TC.PutValue('PE3_SALARIE', Salarie);
              TC.PutValue('PE3_NOM', Nom);
              TC.PutValue('PE3_PRENOM', Prenom);
              TC.PutValue('PE3_DATENAISSANCE', DateNaiss);
              TC.PutValue('PE3_MSAACTIVITE', ActiviteMsa);
              TC.PutValue('PE3_NUMEROSS', NumSS);
              TC.PutValue('PE3_ETABLISSEMENT', TobPaie.Detail[i].GetValue('PPU_ETABLISSEMENT'));
              TC.PutValue('PE3_DATEDEBUT', TobPaie.Detail[i].GetValue('PPU_DATEDEBUT'));
              TC.PutValue('PE3_DATEFIN', TobPaie.Detail[i].GetValue('PPU_DATEFIN'));
              TC.PutValue('PE3_ORDRE', 0);
              TC.PutValue('PE3_UGMSA', UGMsa);
              TC.PutValue('PE3_TOPPLAF', '-');
              TC.PutValue('PE3_PEXOMSA', '');  //PT15
              TC.PutValue('PE3_NBJOURS', Jours);
              TC.PutValue('PE3_NBHEURES', NbHeures);
              if NumElement = 5 then
              begin
                TC.InsertorUpdateDB(False);
                TC.PutValue('PE3_ORDRE', 1);
                NumElement := 0;
              end;
              NumElement := NumElement + 1;
              TC.PutValue('PE3_ELTCALCMSA' + IntToStr(NumElement), '25');
              TC.PutValue('PE3_MONTANTELT' + IntToStr(NumElement), Q.FindField('PHB_MTSALARIAL').AsFloat);
              for x := NumElement + 1 to 5 do
              begin
                TC.PutValue('PE3_ELTCALCMSA' + IntToStr(x), '');
                TC.PutValue('PE3_MONTANTELT' + IntToStr(x), 0);
              end;
              TC.InsertorUpdateDB(False);
            end;
            Ferme(Q);
          end;
          if NumElement = 0 then Writeln(FRapport, '- Aucune rémunération pour la période du ' + DateToStr(debPaie) + ' au ' + DateToStr(FinPaie) + ' est égale à 0');
          //FIN PT22

          if MajEvolutions = True then //Si recherche évolutions du salarié
          begin
            //Enregistrement PE21 Dates entrée et sortie
            DateEntree := TS.GetValue('PSA_DATEENTREE');
            DateSortie := TS.GetValue('PSA_DATESORTIE');
            if (DateEntree >= DateDebut) and (DateEntree <= DateFin) then
            begin
              St := 'SELECT PE2_SALARIE FROM MSAEVOLUTIONSPE2 WHERE ' +
                'PE2_TYPEEVOLMSA="PE21" AND PE2_SALARIE="' + Salarie + '" AND PE2_DEBACTIVITE="' + UsDateTime(DateEntree) + '"';
              TobEvolutions := Tob.Create('MSAEVOLUTIONSPE2', nil, -1);
              TobEvolutions.LoadDetailDBFromSQL ('MSAEVOLUTIONSPE2', St);
              TE := TobEvolutions.FindFirst(['PE2_SALARIE'], [Salarie], False);
              if TE = nil then
              begin
                TE := Tob.Create('MSAEVOLUTIONSPE2', TobEvolutions, -1);
                TE.PutValue('PE2_ETABLISSEMENT', EtablissementS);
                TE.PutValue('PE2_SALARIE', Salarie);
                TE.PutValue('PE2_NOM', Nom);
                TE.PutValue('PE2_PRENOM', Prenom);
                TE.PutValue('PE2_DATENAISSANCE', DateNaiss);
                TE.PutValue('PE2_MSAACTIVITE', ActiviteMsa); //PT10
                TE.PutValue('PE2_NUMEROSS', NumSS);
                TE.PutValue('PE2_TYPEEVOLMSA', 'PE21');
                TE.PutValue('PE2_DATEEFFET', DateEntree);
                TE.PutValue('PE2_DEBACTIVITE', DateEntree);
                If DateEntree = DateSortie then TE.PutValue('PE2_FINACTIVITE', DateSortie); //PT8
                TE.InsertDB(nil, False);
              end;
              FreeAndNil(TobEvolutions);
            end;
            if (DateSortie >= DateDebut) and (DateSortie <= DateFin) and (Datesortie <> dateEntree) then //PT8
            begin
              St := 'SELECT PE2_SALARIE FROM MSAEVOLUTIONSPE2 WHERE ' +
                'PE2_TYPEEVOLMSA="PE21" AND PE2_SALARIE="' + Salarie + '" AND PE2_FINACTIVITE="' + UsDateTime(DateSortie) + '"';
              TobEvolutions := Tob.Create('MSAEVOLUTIONSPE2', nil, -1);
              TobEvolutions.LoadDetailDBFromSQL ('MSAEVOLUTIONSPE2', St);
              TE := TobEvolutions.FindFirst(['PE2_SALARIE'], [Salarie], False);
              if TE = nil then
              begin
                TE := Tob.Create('MSAEVOLUTIONSPE2', TobEvolutions, -1);
                TE.PutValue('PE2_ETABLISSEMENT', EtablissementS);
                TE.PutValue('PE2_SALARIE', Salarie);
                TE.PutValue('PE2_NOM', Nom);
                TE.PutValue('PE2_PRENOM', Prenom);
                TE.PutValue('PE2_DATENAISSANCE', DateNaiss);
                TE.PutValue('PE2_MSAACTIVITE', ActiviteMsa); //PT10
                TE.PutValue('PE2_NUMEROSS', NumSS);
                TE.PutValue('PE2_TYPEEVOLMSA', 'PE21');
                TE.PutValue('PE2_DATEEFFET', DateSortie);
                TE.PutValue('PE2_FINACTIVITE', DateSortie);
                TE.InsertDB(nil, False);
              end;
              FreeAndNil(TobEvolutions);
            end;
          end;
          Writeln(FRapport, '');
        end;

        //DEBUT PT14-1
        St := 'SELECT * FROM ABSENCESALARIE'+
         ' LEFT JOIN MOTIFABSENCE ON ##PMA_PREDEFINI##'+
        ' PCN_TYPECONGE=PMA_MOTIFABSENCE '+
        'WHERE PCN_SALARIE="'+Salarie+'" AND (PMA_TYPEABS="MAN" OR PMA_TYPEABS="MAP") AND '+
        'PCN_DATEDEBUTABS>="'+UsDateTime(DateDebut)+'" AND PCN_DATEDEBUTABS<="'+UsDateTime(DateFin)+'"';
        TobAbs := Tob.Create('les absences',Nil,-1);
        TobAbs.LoadDetailDBFromSQL ('les absences', St);
        For NbAbs := 0 to TobAbs.Detail.Count - 1 do
        begin
          St := 'SELECT PE2_SALARIE FROM MSAEVOLUTIONSPE2 WHERE ' +
            'PE2_TYPEEVOLMSA="PE22" AND PE2_DATEEFFET="'+UsDateTime(TobAbs.Detail[NbAbs].GetValue('PCN_DATEDEBUTABS'))+'" AND '+
            'PE2_SALARIE="' + Salarie + '" AND PE2_FINACTIVITE="' + UsDateTime(TS.GetValue('PSA_DATESORTIE')) + '"';
          TobEvolutions := Tob.Create('MSAEVOLUTIONSPE2', nil, -1);
          TobEvolutions.LoadDetailDBFromSQL ('MSAEVOLUTIONSPE2', St);
          TE := TobEvolutions.FindFirst(['PE2_SALARIE'], [Salarie], False);
          if TE = nil then TE := Tob.Create('MSAEVOLUTIONSPE2', TobEvolutions, -1);
          TE.PutValue('PE2_ETABLISSEMENT', EtablissementS);
          TE.PutValue('PE2_SALARIE', Salarie);
          TE.PutValue('PE2_NOM', Nom);
          TE.PutValue('PE2_PRENOM', Prenom);
          TE.PutValue('PE2_DATENAISSANCE', DateNaiss);
          TE.PutValue('PE2_MSAACTIVITE', ActiviteMsa);
          TE.PutValue('PE2_NUMEROSS', NumSS);
          TE.PutValue('PE2_TYPEEVOLMSA', 'PE22');
          TE.PutValue('PE2_DATEEFFET', TobAbs.Detail[NbAbs].GetValue('PCN_DATEDEBUTABS'));
          TE.PutValue('PE2_DEBSUSPCT', TobAbs.Detail[NbAbs].GetValue('PCN_DATEDEBUTABS'));
          TE.PutValue('PE2_FINSUSSPCT', TobAbs.Detail[NbAbs].GetValue('PCN_DATEFINABS'));
          TE.PutValue('PE2_MSASUSPCT', '06');
          TE.InsertOrUpdateDB(False);
          FreeAndNil(TobEvolutions);
        end;
        FreeAndNil(TobAbs);
        //FIN PT14-1
        FreeAndNil(TobPaie);
        FreeAndNil(TobMSAPeriodes);
        MoveCurProgressForm('Salarié ' + Salarie + ' ' + Nom + ' ' + Prenom);
        Q_Mul.Next;
      end;
      FiniMoveProgressForm;
    end;
    FreeAndNil(TobSalaries); //PT14
    Writeln(FRapport, 'Préparation déclaration MSA terminée : ' + DateTimeToStr(Now));
    CloseFile(FRapport);
    i := PGIAsk('Voulez-vous visualiser le fichier de contrôle ?',
      'Calcul des périodes déclarations MSA');

    if i = mrYes then
    begin
      {$IFDEF EAGLCLIENT}
      NomFicRapport := '"'+VH_PAIE.PGCheminEagl + '\' + 'MSA_PGI.log"';
      {$ENDIF}
      ShellExecute(0, PCHAR('open'), PChar('WordPad'), PChar(NomFicRapport), nil, SW_RESTORE);
    end;
    TFMul(Ecran).BCherche.Click;
    CommitTrans;
  except
    Rollback;
    CloseFile(FRapport);
  end;
end;

procedure TOF_PGMSACALCPERIODES.DateElipsisClick(Sender: TObject);
var
  key: char;
begin
  key := '*';
  ParamDate(Ecran, Sender, Key);
end;

procedure TOF_PGMSACALCPERIODES.ExitEdit(Sender: TObject);
var
  edit: thedit;
begin
  edit := THEdit(Sender);
  if edit <> nil then //AffectDefautCode que si gestion du code salarié en Numérique
    if (VH_Paie.PgTypeNumSal = 'NUM') and (length(Edit.text) < 11) and (isnumeric(edit.text)) then
      edit.text := AffectDefautCode(edit, 10);
end;

Function TOF_PGMSACALCPERIODES.ControlActivite : Boolean;
var TobAct : Tob;
//    Q : TQUery;
    Abrege,Lib,Code,MessError : String;
    i,NbError : Integer;
    St : String;
begin
     Result := True;
//Optimisation     Q := OpenSQL('SELECT * FROM CHOIXCOD WHERE CC_TYPE="PMS"',True);
//     TobAct.LoadDetailDB('LesACtivites','','',Q,False);
     St := 'SELECT * FROM CHOIXCOD WHERE CC_TYPE="PMS"';
     TobAct := Tob.Create('Table',Nil,-1);
     TobAct.LoadDetailDBFromSQL ('LesACtivites', St);
//     Ferme(Q);
     MessError := '';
     NbError := 0;
     For i := 0 to TobAct.Detail.Count - 1 do
     begin
          Abrege := TobAct.Detail[i].GetValue('CC_ABREGE');
          Lib := TobAct.Detail[i].GetValue('CC_LIBELLE');
          Code := TobAct.Detail[i].GetValue('CC_CODE');
          If (Length(Abrege) <> 4) or (Not IsNumeric(Abrege)) then
          begin
               MessError := '#13#10 - ' + Code + ' ' + Lib;
               NbError := Nberror + 1;
          end;
     end;
     FreeAndNil(TobAct);
     If MessError <> '' then
     begin
          If NbError > 1 then PGIBox('Le calcul ne peut être lancé car l''abrégé des activités suivantes doit être numérique et sur  4 caractères :'+MessError)
          else PGIBox('Le calcul ne peut être lancé car l''abrégé de l''activité suivante doit être numérique et sur  4 caractères :'+MessError);
          Result := False;
     end;
end;

initialization
  registerclasses([TOF_PGMSACALCPERIODES]);
end.

