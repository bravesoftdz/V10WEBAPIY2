{***********UNITE*************************************************
Auteur  ...... : JL
Créé le ...... : 28/03/2003
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : PGCALCCONGESSPEC ()
Mots clefs ... : TOF;PGCALCCONGESSPEC
*****************************************************************
PT1 09/02/2006 V_650 JL FQ 12907 Correction pour numéro congés spectacle erroné a partir 2eme période
PT2 22/09/2006 V_700 JL FQ 13402 pour le calcul de BASECONGE reprendre cumul 01 brut avant abattement
PT3 18/07/2007 V_72  LF FQ 14588 : on encadre le nom du fichier (Chemin + Fichier) par des guillemets
PT4 03/12/2007 V_72  RM FQ 14020 : Compléter le matricule salarié par des zeros automatiquement
}
unit UTofPGCalcCongesSpec;

interface

uses StdCtrls, Controls, Classes,
  {$IFNDEF EAGLCLIENT}
  db, {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF} Mul, Fe_Main, HDB, DBGrids,
  {$ELSE}
  eMul, MainEAGL,
  {$ENDIF}
  forms, sysutils, ComCtrls, HCtrls, HEnt1, HMsgBox, UTOF, UTob, HTB97, PG_OutilsEnvoi, HStatus,
  EntPaie, AGLInit, ed_tools, ParamDat, HQry, ShellAPI, windows, PGOutils,PGOutils2;

type
  TOF_PGCALCCONGESSPEC = class(TOF)
    procedure OnArgument(S: string); override;
  private
    procedure CalculPeriode(Sender: TObject);
    procedure DateElipsisClick(Sender: TObject);
    Procedure ExitEdit(Sender: TObject);  //PT4
  end;

implementation

procedure TOF_PGCALCCONGESSPEC.OnArgument(S: string);
var
  DateDebut, DateFin: TDateTime;
  BCalcul: TToolBarButton97;
  THDate: THEdit;
  Arg: string;
  aa, mm, jj, Mois: word;
  Defaut : THEdit;  //PT4
begin
  inherited;
  Arg := ReadTokenPipe(S, ';');
  BCalcul := TToolBarButton97(GetControl('BCALCULER'));
  if BCalcul <> nil then BCalcul.OnClick := CalculPeriode;
  DecodeDate(Date, aa, mm, jj);
  Mois := 1;
  if (mm < 7) and (mm > 3) then Mois := 4;
  if (mm < 10) and (mm > 6) then Mois := 7;
  if mm > 9 then Mois := 10;
  DateDebut := EncodeDate(aa, Mois, 1);
  DateFin := PlusMois(DateDebut, 2);
  DateFin := FinDeMois(DateFin);
  SetControlText('DATEDEBUT', DateToStr(Datedebut));
  SetControltext('DATEFIN', dateToStr(DateFin));
  THDate := THEdit(GetControl('DATEDEBUT'));
  if THDate <> nil then THDate.OnElipsisClick := DateElipsisClick;
  THDate := THEdit(GetControl('DATEFIN'));
  if THDate <> nil then THDate.OnElipsisClick := DateElipsisClick;
  SetControlCaption('LIBSAL', '');
  SetControltext('PSA_ETABLISSEMENT', '');

  Defaut:=ThEdit(getcontrol('PSA_SALARIE'));      //PT4
  If Defaut<>nil then Defaut.OnExit:=ExitEdit;    //PT4
end;

procedure TOF_PGCALCCONGESSPEC.CalculPeriode(Sender: TObject);
var
  TobCongesSpec, TobPaie, TC, TVerif, TobCreation, TobCumulSS, TobCumulBrut, TCS, TCB: Tob;
  Q: TQuery;
  NumSS, CleSS, NumCS, NumCSCopy, CleCS, NomNaiss, NomEpoux, Prenom, Pseudo, Salarie, LibSalarie, Sexe: string;
  Emploi, Cadre, JJd, MMd, AAd, Datedebut, DateFin, JJf, MMf, AAf: string;
  PayeLe, aa, mm, jj: string;
  Longueur, i, j: Integer;
  verifSal, NomFicRapport: string;
  VerifDateDebut, VerifDateFin: TDateTime;
  BaseConge, Salaire, NbJoursCachets: Double;
  {$IFNDEF EAGLCLIENT}
  Liste: THDBGrid;
  {$ELSE}
  Liste: THGrid;
  {$ENDIF}
  Q_Mul: THQuery;
  FRapport: TextFile;
begin
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
  NomFicRapport := VH_Paie.PgCheminEagl + '\CONGESSPEC_PGI.log';
  {$ELSE}
  NomFicRapport := V_PGI.DatPath + '\' + 'CONGESSPEC_PGI.log';
  {$ENDIF}

  if FileExists(NomFicRapport) then DeleteFile(PChar(NomFicRapport));
  AssignFile(FRapport, NomFicRapport);
  {$I-}ReWrite(FRapport);
  {$I+}
  if IoResult <> 0 then
  begin
    PGIBox('Fichier inaccessible : ' + NomFicRapport,
      'Rapport calcul congés spectacles');
  end;
  try
    begintrans;
    Writeln(FRapport, '');
    Writeln(FRapport, 'Début du calcul des congés spectacles : ' + DateTimeToStr(Now));
    Writeln(FRapport, 'Attention, les zones signataire, Date délivrance, et lieu seront renseignées lors de l''édition des certificats d''emploi');
    {$IFDEF EAGLCLIENT}
    if (TFMul(Ecran).bSelectAll.Down) then
      TFMul(Ecran).Fetchlestous;
    {$ENDIF}
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
        NumSS := '';
        CleSS := '';
        NomNaiss := '';
        NomEpoux := '';
        Prenom := '';
        Pseudo := '';
        Salarie := TFmul(Ecran).Q.FindField('PSE_SALARIE').asstring;
        Q := OpenSQL('SELECT PSA_PRENOM,PSA_LIBELLE,PSA_NOMJF,PSA_SURNOM,PSA_DADSCAT,PSA_SEXE' +
          ',PSA_NUMEROSS,PSE_ISCONGSPEC FROM SALARIES' +
          '  LEFT JOIN DEPORTSAL ON PSA_SALARIE=PSE_SALARIE' +
          ' WHERE PSA_SALARIE="' + Salarie + '"', True);
        if not Q.Eof then
        begin
          NumSS := Q.FindField('PSA_NUMEROSS').AsString;
          Longueur := Length(NumSS);
          if Longueur > 13 then CleSS := Copy(NumSS, 14, 2)
          else CleSS := '';
          if NumSS <> '' then SetLength(NumSS, 13);
          LibSalarie := Q.FindField('PSA_LIBELLE').AsString;
          if Q.FindField('PSA_NOMJF').AsString <> '' then
          begin
            NomNaiss := Q.FindField('PSA_NOMJF').AsString;
            NomEpoux := Q.FindField('PSA_LIBELLE').AsString;
            if Length(NomNaiss) > 20 then SetLength(NomNaiss, 20);
            if Length(NomEpoux) > 20 then SetLength(NomEpoux, 20);
          end
          else
          begin
            NomNaiss := Q.FindField('PSA_LIBELLE').AsString;
            NomEpoux := '';
            if Length(NomNaiss) > 20 then SetLength(NomNaiss, 20);
          end;
          Prenom := Q.FindField('PSA_PRENOM').AsString;
          if Length(Prenom) > 12 then SetLength(Prenom, 12);
          Pseudo := Q.FindField('PSA_SURNOM').AsString;
          if Length(Pseudo) > 20 then SetLength(Pseudo, 20);
          if (Q.FindField('PSA_DADSCAT').AsString = '01') or (Q.FindField('PSA_DADSCAT').AsString = '02') then
            Cadre := 'O'
          else Cadre := 'N';
          Sexe := Q.FindField('PSA_SEXE').AsString;
          NumCS := Q.FindField('PSE_ISCONGSPEC').AsString;
        end;
        Ferme(Q);
        if NumCS = '' then Writeln(FRapport, '- le numéro d''identification aux congés spectacles n''est pas renseigné');
        if TestNumeroSS(NumSS + CleSS, Sexe) < 0 then Writeln(FRapport, '- Vérifier le N° de sécurité sociale');
        Q := OpenSQL('SELECT PPU_SALARIE,PPU_LIBELLEEMPLOI,PPU_DATEDEBUT,PPU_DATEFIN,PPU_PAYELE,PPU_CBRUT,PPU_CBASESS' +
          ' FROM PAIEENCOURS' +
          ' WHERE PPU_SALARIE="' + Salarie + '" AND' +
          ' PPU_DATEDEBUT>="' + UsDatetime(StrToDate(GetControlText('DATEDEBUT'))) + '"' +
          ' AND PPU_DATEFIN<="' + UsDatetime(StrToDate(GetControlText('DATEFIN'))) + '"', True);
        TobPaie := Tob.Create('Les paies', nil, -1);
        TobPaie.LoadDetailDB('PAIEENCOURS', '', '', Q, False);
        Ferme(Q);
        Q := OpenSQL('SELECT SUM (PHC_MONTANT) BRUT FROM HISTOCUMSAL LEFT JOIN DEPORTSAL ON PHC_SALARIE=PSE_SALARIE ' +
          'WHERE PHC_DATEDEBUT>="' + UsDatetime(StrToDate(GetControlText('DATEDEBUT'))) + '"' +
          ' AND PHC_DATEFIN<="' + UsDatetime(StrToDate(GetControlText('DATEFIN'))) + '"' +
          ' AND PHC_CUMULPAIE="01"' +
          ' AND PHC_SALARIE="' + Salarie + '"' +
          ' GROUP BY PHC_DATEDEBUT,PHC_DATEFIN,PHC_SALARIE', True);
        TobCumulBrut := Tob.Create('Les cumuls', nil, -1);
        TobCumulBrut.LoadDetailDB('Les cumuls', '', '', Q, False);
        Ferme(Q);
        Q := OpenSQL('SELECT SUM (PHC_MONTANT) BASESS FROM HISTOCUMSAL LEFT JOIN DEPORTSAL ON PHC_SALARIE=PSE_SALARIE ' +
          'WHERE PHC_DATEDEBUT>="' + UsDatetime(StrToDate(GetControlText('DATEDEBUT'))) + '"' +
          ' AND PHC_DATEFIN<="' + UsDatetime(StrToDate(GetControlText('DATEFIN'))) + '"' +
          ' AND PHC_CUMULPAIE="01"' +//PT2
          ' AND PHC_SALARIE="' + Salarie + '"' +
          ' GROUP BY PHC_DATEDEBUT,PHC_DATEFIN,PHC_SALARIE', True);
        TobCumulSS := Tob.Create('Les cumuls', nil, -1);
        TobCumulSS.LoadDetailDB('Les cumuls', '', '', Q, False);
        Ferme(Q);
        Q := OpenSQL('SELECT * FROM CONGESSPEC WHERE PCS_SALARIE="' + Salarie + '"' +
          ' AND PCS_DATEDEBUT>="' + UsDatetime(StrToDate(GetControlText('DATEDEBUT'))) + '"' +
          ' AND PCS_DATEFIN<="' + UsDatetime(StrToDate(GetControlText('DATEFIN'))) + '"', True);
        TobCongesSpec := Tob.create('CONGESSPEC', nil, -1);
        TobCongesSpec.LoadDetailDB('CONGESSPEC', '', '', Q, False);
        Ferme(Q);
        Writeln(FRapport, '     Salarie :' + Salarie + ' ' + LibSalarie + ' ' + Prenom);
        Writeln(FRapport, '');
        if TobPaie.Detail.Count <= 0 then Writeln(FRapport, 'Ce salarié n''a aucune période de paie');
        for i := 0 to TobPaie.Detail.Count - 1 do
        begin
          TC := Tob.Create('CONGESSPEC', TobCongesSpec, -1);
          TC.PutValue('PCS_SALARIE', TobPaie.Detail[i].GetValue('PPU_SALARIE'));
          VerifSal := TobPaie.Detail[i].GetValue('PPU_SALARIE');
          TC.PutValue('PCS_DATEDEBUT', TobPaie.Detail[i].GetValue('PPU_DATEDEBUT'));
          VerifDateDebut := TobPaie.Detail[i].GetValue('PPU_DATEDEBUT');
          TC.PutValue('PCS_DATEFIN', TobPaie.Detail[i].GetValue('PPU_DATEFIN'));
          VerifDateFin := TobPaie.Detail[i].GetValue('PPU_DATEFIN');
          Writeln(FRapport, '********************** Période du ' + dateToStr(VerifDateDebut) + ' au ' + DateToStr(VerifDateFin));
          TVerif := TobCongesSpec.FindFirst(['PCS_SALARIE', 'PCS_DATEDEBUT', 'PCS_DATEFIN'], [VerifSal, VerifDateDebut, VerifDateFin], False);
          if TVerif <> nil then TVerif.DeleteDB(False);
          TC.PutValue('PCS_NUMEROSS', NumSS);
          TC.PutValue('PCS_CLENUMSS', CleSS);
          Longueur := Length(NumCS);
          if Longueur > 1 then CleCS := NumCS[1]
          else CleCS := '';
          TC.PutValue('PCS_CLENUMCONG', CleCS);
          if Longueur > 1 then NumCSCopy := Copy(NumCS, 2, Longueur - 1)    //PT1
          else NumCSCopy := '';
          TC.PutValue('PCS_NUMCONGES', NumCSCopy);
          TC.PutValue('PCS_NOMNAISS', NomNaiss);
          TC.PutValue('PCS_NOMEPOUX', NomEpoux);
          TC.PutValue('PCS_PRENOM', Prenom);
          TC.PutValue('PCS_PSEUDO', Pseudo);
          Emploi := TobPaie.Detail[i].GetValue('PPU_LIBELLEEMPLOI');
          if Length(Emploi) > 15 then SetLength(Emploi, 15);
          TC.PutValue('PCS_EMPLOI', Emploi);
          TC.PutValue('PCS_CADRE', Cadre);
          DateDebut := DateToStr(TobPaie.Detail[i].GetValue('PPU_DATEDEBUT'));
          DateFin := DateToStr(TobPaie.Detail[i].GetValue('PPU_DATEFIN'));
          JJd := ReadTokenPipe(DateDebut, '/');
          MMd := ReadTokenPipe(DateDebut, '/');
          AAd := ReadTokenPipe(DateDebut, '/');
          JJf := ReadTokenPipe(DateFin, '/');
          MMf := ReadTokenPipe(DateFin, '/');
          AAf := ReadTokenPipe(DateFin, '/');
          TC.PutValue('PCS_DATESTRAVAIL', JJd + MMd + AAd + JJf + MMf + AAf);
          Q := OpenSQL('SELECT PCI_ISNBEFFECTUE FROM CONTRATTRAVAIL ' +
            'WHERE PCI_SALARIE="' + TobPaie.Detail[i].GetValue('PPU_SALARIE') + '" ' +
            'AND PCI_DEBUTCONTRAT="' + UsDateTime(TobPaie.Detail[i].GetValue('PPU_DATEDEBUT')) + '" ' +
            'AND PCI_FINCONTRAT="' + UsDateTime(TobPaie.Detail[i].GetValue('PPU_DATEFIN')) + '" ' +
            'AND PCI_ISCACHETS="X"', True);
          if not Q.Eof then
          begin
            TC.PutValue('PCS_NBJOURS', FloatToStr(Arrondi(Q.FindField('PCI_ISNBEFFECTUE').AsFloat, 0)));
            Writeln(FRapport, '- Le salarié possède un contrat et est rémunéré en nombre de cachets');
            Ferme(Q);
          end
          else
          begin
            Ferme(Q);
            Q := OpenSQL('SELECT PHB_MTREM FROM HISTOBULLETIN WHERE PHB_SALARIE="' + Salarie + '"' +
              ' AND PHB_DATEDEBUT="' + UsdateTime(VerifDateDebut) + '" AND PHB_DATEFIN="' + UsDateTime(VerifDateFin) + '"' +
              ' AND PHB_NATURERUB="AAA" AND PHB_RUBRIQUE="0280"', True);
            if not Q.Eof then
            begin
              if Q.FindField('PHB_MTREM').AsFloat > 0 then TC.PutValue('PCS_NBJOURS', FloatToStr(Q.FindField('PHB_MTREM').AsFloat))
              else
              begin
                TC.PutValue('PCS_NBJOURS', FloatToStr(Arrondi(VerifDateFin - VerifDateDebut, 0) + 1));
                Writeln(FRapport, '- Aucun contrat pour ce salarié, le nombre de jours par défaut est de ' + FloatToSTr(VerifDateFin - VerifDateDebut + 1));
              end;
            end
            else
            begin
              TC.PutValue('PCS_NBJOURS', FloatToStr(Arrondi(VerifDateFin - VerifDateDebut, 0) + 1));
              Writeln(FRapport, '- Aucun contrat pour ce salarié, le nombre de jours par défaut est de ' + FloatToSTr(VerifDateFin - VerifDateDebut + 1));
            end;
            Ferme(Q);
          end;
          TCB := TobCumulBrut.FindFirst(['PHC_SALARIE', 'PHC_DATEDEBUT', 'PHC_DATEFIN'], [VerifSal, VerifDateDebut, VerifDateFin], False);
          if TCB <> nil then Salaire := Arrondi(TCB.GetValue('BRUT'), 0)
          else Salaire := Arrondi(TobPaie.Detail[i].GetValue('PPU_CBRUT'), 0);
          TCS := TobCumulSS.FindFirst(['PHC_SALARIE', 'PHC_DATEDEBUT', 'PHC_DATEFIN'], [VerifSal, VerifDateDebut, VerifDateFin], False);
          if TCS <> nil then BaseConge := Arrondi(TCS.GetValue('BASESS'), 0)
          else BaseConge := Arrondi(TobPaie.Detail[i].GetValue('PPU_CBASESS'), 0);
          if BaseConge < 0 then BaseConge := 0;
          if Salaire < 0 then Salaire := 0;
          TC.PutValue('PCS_BASECONGE', FloatToStr(BaseConge));
          if BaseConge = 0 then Writeln(FRapport, '- La base congés est à 0');
          if Salaire = 0 then Writeln(FRapport, '- La salaire brut est à 0');
          TC.PutValue('PCS_SALBRUT', FloatToStr(Salaire));
          if TobPaie.Detail[i].GetValue('PPU_PAYELE') <= IDate1900 then
          begin
            TC.PutValue('PCS_DATEPAIEM', '');
            Writeln(FRapport, '- La date de paiement n''est pas renseignée');
          end
          else
          begin
            PayeLe := DateToStr(TobPaie.Detail[i].GetValue('PPU_PAYELE'));
            jj := ReadTokenPipe(PayeLe, '/');
            mm := ReadTokenPipe(PayeLe, '/');
            aa := ReadTokenPipe(PayeLe, '/');
            TC.PutValue('PCS_DATEPAIEM', jj + mm + aa);
          end;
          TC.PutValue('PCS_LIEU', '');
          TC.PutValue('PCS_DATEDELIV', '');
          TC.PutValue('PCS_NOMSIGN', '');
          TC.PutValue('PCS_NUMCERTIFICAT', '');
          TC.PutValue('PCS_NUMENTREP', '');
          TC.PutValue('PCS_LETTRECLE', '');
          TC.PutValue('PCS_LIBRE', '');
          TC.InsertorUpdateDB(False);
          Writeln(FRapport, '');
          MoveCurProgressForm(LibSalarie);
        end;
        TobPaie.Free;
        TobCongesSpec.Free;
        TobCumulBrut.Free;
        TobCumulSS.Free;
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
        NumSS := '';
        CleSS := '';
        NomNaiss := '';
        NomEpoux := '';
        Prenom := '';
        Pseudo := '';
        Salarie := TFmul(Ecran).Q.FindField('PSE_SALARIE').AsString;
        Q := OpenSQL('SELECT PSA_PRENOM,PSA_LIBELLE,PSA_NOMJF,PSA_SURNOM,PSA_DADSCAT,PSA_SEXE' +
          ',PSA_NUMEROSS,PSE_ISCONGSPEC FROM SALARIES' +
          '  LEFT JOIN DEPORTSAL ON PSA_SALARIE=PSE_SALARIE' +
          ' WHERE PSA_SALARIE="' + Salarie + '"', True);
        if not Q.Eof then
        begin
          NumSS := Q.FindField('PSA_NUMEROSS').AsString;
          Longueur := Length(NumSS);
          if Longueur > 13 then CleSS := Copy(NumSS, 14, 2)
          else CleSS := '';
          if NumSS <> '' then SetLength(NumSS, 13);
          LibSalarie := Q.FindField('PSA_LIBELLE').AsString;
          if Q.FindField('PSA_NOMJF').AsString <> '' then
          begin
            NomNaiss := Q.FindField('PSA_NOMJF').AsString;
            NomEpoux := Q.FindField('PSA_LIBELLE').AsString;
            if Length(NomNaiss) > 20 then SetLength(NomNaiss, 20);
            if Length(NomEpoux) > 20 then SetLength(NomEpoux, 20);
          end
          else
          begin
            NomNaiss := Q.FindField('PSA_LIBELLE').AsString;
            NomEpoux := '';
            if Length(NomNaiss) > 20 then SetLength(NomNaiss, 20);
          end;
          Prenom := Q.FindField('PSA_PRENOM').AsString;
          if Length(Prenom) > 12 then SetLength(Prenom, 12);
          Pseudo := Q.FindField('PSA_SURNOM').AsString;
          if Length(Pseudo) > 20 then SetLength(Pseudo, 20);
          if (Q.FindField('PSA_DADSCAT').AsString = '01') or (Q.FindField('PSA_DADSCAT').AsString = '02') then
            Cadre := 'O'
          else Cadre := 'N';
          Sexe := Q.FindField('PSA_SEXE').AsString;
          NumCS := Q.FindField('PSE_ISCONGSPEC').AsString;
        end;
        Ferme(Q);
        if NumCS = '' then Writeln(FRapport, '- le numéro d''identification aux congés spectacles n''est pas renseigné');
        if TestNumeroSS(NumSS + CleSS, Sexe) > 0 then Writeln(FRapport, '- Vérifier le N° de sécurité sociale');
        Q := OpenSQL('SELECT PPU_SALARIE,PPU_LIBELLEEMPLOI,PPU_DATEDEBUT,PPU_DATEFIN,PPU_PAYELE,PPU_CBRUT,PPU_CBASESS' +
          ' FROM PAIEENCOURS' +
          ' WHERE PPU_SALARIE="' + Salarie + '" AND' +
          ' PPU_DATEDEBUT>="' + UsDatetime(StrToDate(GetControlText('DATEDEBUT'))) + '"' +
          ' AND PPU_DATEFIN<="' + UsDatetime(StrToDate(GetControlText('DATEFIN'))) + '"', True);
        TobPaie := Tob.Create('Les paies', nil, -1);
        TobPaie.LoadDetailDB('PAIEENCOURS', '', '' ,Q, False);
        Ferme(Q);
        Q := OpenSQL('SELECT SUM (PHC_MONTANT) BRUT FROM HISTOCUMSAL LEFT JOIN DEPORTSAL ON PHC_SALARIE=PSE_SALARIE ' +
          'WHERE PHC_DATEDEBUT>="' + UsDatetime(StrToDate(GetControlText('DATEDEBUT'))) + '"' +
          ' AND PHC_DATEFIN<="' + UsDatetime(StrToDate(GetControlText('DATEFIN'))) + '"' +
          ' AND PHC_CUMULPAIE="01"' +
          ' AND PHC_SALARIE="' + Salarie + '"' +
          ' GROUP BY PHC_DATEDEBUT,PHC_DATEFIN,PHC_SALARIE', True);
        TobCumulBrut := Tob.Create('Les cumuls', nil, -1);
        TobCumulBrut.LoadDetailDB('Les cumuls', '', '', Q, False);
        Ferme(Q);
        Q := OpenSQL('SELECT SUM (PHC_MONTANT) BASESS FROM HISTOCUMSAL LEFT JOIN DEPORTSAL ON PHC_SALARIE=PSE_SALARIE ' +
          'WHERE PHC_DATEDEBUT>="' + UsDatetime(StrToDate(GetControlText('DATEDEBUT'))) + '"' +
          ' AND PHC_DATEFIN<="' + UsDatetime(StrToDate(GetControlText('DATEFIN'))) + '"' +
          ' AND PHC_CUMULPAIE="01"' + //PT2
          ' AND PHC_SALARIE="' + Salarie + '"' +
          ' GROUP BY PHC_DATEDEBUT,PHC_DATEFIN,PHC_SALARIE', True);
        TobCumulSS := Tob.Create('Les cumuls', nil, -1);
        TobCumulSS.LoadDetailDB('Les cumuls', '', '', Q, False);
        Ferme(Q);
        Q := OpenSQL('SELECT * FROM CONGESSPEC WHERE PCS_SALARIE="' + Salarie + '"' +
          ' AND PCS_DATEDEBUT>="' + UsDatetime(StrToDate(GetControlText('DATEDEBUT'))) + '"' +
          ' AND PCS_DATEFIN<="' + UsDatetime(StrToDate(GetControlText('DATEFIN'))) + '"', True);
        TobCongesSpec := Tob.create('CONGESSPEC', nil, -1);
        TobCongesSpec.LoadDetailDB('CONGESSPEC', '', '', Q, False);
        Ferme(Q);
        Writeln(FRapport, '     Salarie :' + Salarie + ' ' + LibSalarie + ' ' + Prenom);
        Writeln(FRapport, '');
        if TobPaie.Detail.Count <= 0 then Writeln(FRapport, 'Ce salarié n''a aucune période de paie');
        for i := 0 to TobPaie.Detail.Count - 1 do
        begin
          TC := Tob.Create('CONGESSPEC', TobCongesSpec, -1);
          TC.PutValue('PCS_SALARIE', TobPaie.Detail[i].GetValue('PPU_SALARIE'));
          VerifSal := TobPaie.Detail[i].GetValue('PPU_SALARIE');
          TC.PutValue('PCS_DATEDEBUT', TobPaie.Detail[i].GetValue('PPU_DATEDEBUT'));
          VerifDateDebut := TobPaie.Detail[i].GetValue('PPU_DATEDEBUT');
          TC.PutValue('PCS_DATEFIN', TobPaie.Detail[i].GetValue('PPU_DATEFIN'));
          VerifDateFin := TobPaie.Detail[i].GetValue('PPU_DATEFIN');
          Writeln(FRapport, '********************** Période du ' + dateToStr(VerifDateDebut) + ' au ' + DateToStr(VerifDateFin));
          TVerif := TobCongesSpec.FindFirst(['PCS_SALARIE', 'PCS_DATEDEBUT', 'PCS_DATEFIN'], [VerifSal, VerifDateDebut, VerifDateFin], False);
          if TVerif <> nil then TVerif.DeleteDB(False);
          TC.PutValue('PCS_NUMEROSS', NumSS);
          TC.PutValue('PCS_CLENUMSS', CleSS);
          Longueur := Length(NumCS);
          if Longueur > 1 then CleCS := NumCS[1]
          else CleCS := '';
          TC.PutValue('PCS_CLENUMCONG', CleCS);
          if Longueur > 1 then NumCSCopy := Copy(NumCS, 2, Longueur - 1) //PT1
          else NumCSCopy := '';
          TC.PutValue('PCS_NUMCONGES', NumCSCopy);
          TC.PutValue('PCS_NOMNAISS', NomNaiss);
          TC.PutValue('PCS_NOMEPOUX', NomEpoux);
          TC.PutValue('PCS_PRENOM', Prenom);
          TC.PutValue('PCS_PSEUDO', Pseudo);
          Emploi := TobPaie.Detail[i].GetValue('PPU_LIBELLEEMPLOI');
          if Length(Emploi) > 15 then SetLength(Emploi, 15);
          TC.PutValue('PCS_EMPLOI', Emploi);
          TC.PutValue('PCS_CADRE', Cadre);
          DateDebut := DateToStr(TobPaie.Detail[i].GetValue('PPU_DATEDEBUT'));
          DateFin := DateToStr(TobPaie.Detail[i].GetValue('PPU_DATEFIN'));
          JJd := ReadTokenPipe(DateDebut, '/');
          MMd := ReadTokenPipe(DateDebut, '/');
          AAd := ReadTokenPipe(DateDebut, '/');
          JJf := ReadTokenPipe(DateFin, '/');
          MMf := ReadTokenPipe(DateFin, '/');
          AAf := ReadTokenPipe(DateFin, '/');
          TC.PutValue('PCS_DATESTRAVAIL', JJd + MMd + AAd + JJf + MMf + AAf);
          Q := OpenSQL('SELECT PCI_ISNBEFFECTUE FROM CONTRATTRAVAIL ' +
            'WHERE PCI_SALARIE="' + TobPaie.Detail[i].GetValue('PPU_SALARIE') + '" ' +
            'AND PCI_DEBUTCONTRAT="' + UsDateTime(TobPaie.Detail[i].GetValue('PPU_DATEDEBUT')) + '" ' +
            'AND PCI_FINCONTRAT="' + UsDateTime(TobPaie.Detail[i].GetValue('PPU_DATEFIN')) + '" ' +
            'AND PCI_ISCACHETS="X"', True);
          if not Q.Eof then
          begin
            TC.PutValue('PCS_NBJOURS', FloatToStr(Arrondi(Q.FindField('PCI_ISNBEFFECTUE').AsFloat, 0)));
            Writeln(FRapport, '- Le salarié possède un contrat et est rémunéré en nombre de cachets');
          end
          else
          begin
            TC.PutValue('PCS_NBJOURS', FloatToStr(Arrondi(VerifDateFin - VerifDateDebut, 0) + 1));
            Writeln(FRapport, '- Aucun contrat pour ce salarié, le nombre de jours par défaut est de ' + FloatToSTr(VerifDateFin - VerifDateDebut + 1));
          end;
          Ferme(Q);
          TCB := TobCumulBrut.FindFirst(['PHC_SALARIE', 'PHC_DATEDEBUT', 'PHC_DATEFIN'], [VerifSal, VerifDateDebut, VerifDateFin], False);
          if TCB <> nil then Salaire := Arrondi(TCB.GetValue('BRUT'), 0)
          else Salaire := Arrondi(TobPaie.Detail[i].GetValue('PPU_CBRUT'), 0);
          TCS := TobCumulSS.FindFirst(['PHC_SALARIE', 'PHC_DATEDEBUT', 'PHC_DATEFIN'], [VerifSal, VerifDateDebut, VerifDateFin], False);
          if TCS <> nil then BaseConge := Arrondi(TCS.GetValue('BASESS'), 0)
          else BaseConge := Arrondi(TobPaie.Detail[i].GetValue('PPU_CBASESS'), 0);
          if BaseConge < 0 then BaseConge := 0;
          if Salaire < 0 then Salaire := 0;
          TC.PutValue('PCS_BASECONGE', FloatToStr(BaseConge));
          if BaseConge = 0 then Writeln(FRapport, '- La base congés est à 0');
          if Salaire = 0 then Writeln(FRapport, '- La salaire brut est à 0');
          TC.PutValue('PCS_SALBRUT', FloatToStr(Salaire));
          if TobPaie.Detail[i].GetValue('PPU_PAYELE') <= IDate1900 then
          begin
            TC.PutValue('PCS_DATEPAIEM', '');
            Writeln(FRapport, '- La date de paiement n''est pas renseignée');
          end
          else
          begin
            PayeLe := DateToStr(TobPaie.Detail[i].GetValue('PPU_PAYELE'));
            jj := ReadTokenPipe(PayeLe, '/');
            mm := ReadTokenPipe(PayeLe, '/');
            aa := ReadTokenPipe(PayeLe, '/');
            TC.PutValue('PCS_DATEPAIEM', jj + mm + aa);
          end;
          TC.PutValue('PCS_LIEU', '');
          TC.PutValue('PCS_DATEDELIV', '');
          TC.PutValue('PCS_NOMSIGN', '');
          TC.PutValue('PCS_NUMCERTIFICAT', '');
          TC.PutValue('PCS_NUMENTREP', '');
          TC.PutValue('PCS_LETTRECLE', '');
          TC.PutValue('PCS_LIBRE', '');
          TC.InsertorUpdateDB(False);
          Writeln(FRapport, '');
          MoveCurProgressForm(LibSalarie);
        end;
        TobPaie.Free;
        TobCongesSpec.Free;
        TobCumulBrut.Free;
        TobCumulSS.Free;
        TFmul(Ecran).Q.Next;
      end;
      FiniMoveProgressForm;
    end;
    Writeln(FRapport, 'Préparation des congés spectacles terminée : ' + DateTimeToStr(Now));
    CloseFile(FRapport);
    i := PGIAsk('Voulez-vous visualiser le fichier de contrôle ?',
      'Calcul des congés spectacles');
    if i = mrYes then
// d PT3
    begin
      {$IFDEF EAGLCLIENT}
      NomFicRapport := '"'+VH_Paie.PgCheminEagl + '\CONGESSPEC_PGI.log"';
      {$ENDIF}

      ShellExecute(0, PCHAR('open'), PChar('WordPad'), PChar(NomFicRapport), nil, SW_RESTORE);
    end;
// f PT3
    TFMul(Ecran).BCherche.Click;
    CommitTrans;
  except
    Rollback;
    CloseFile(FRapport);
  end;
end;

procedure TOF_PGCALCCONGESSPEC.DateElipsisClick(Sender: TObject);
var
  key: char;
begin
  key := '*';
  ParamDate(Ecran, Sender, Key);
end;

//PT4 AJout de la procedure ===========================>
Procedure TOF_PGCALCCONGESSPEC.ExitEdit(Sender: TObject);
var edit : thedit;
Begin
edit:=THEdit(Sender);
If edit <> nil then	//AffectDefautCode que si gestion du code salarié en Numérique
    if (VH_Paie.PgTypeNumSal='NUM') and (length(Edit.text)<11) and (isnumeric(edit.text)) then
    edit.text:=AffectDefautCode(edit,10);
End;


initialization
  registerclasses([TOF_PGCALCCONGESSPEC]);
end.

