{***********UNITE*************************************************
Auteur  ...... : PAIE PGI
Créé le ...... : 13/09/2001
Modifié le ... :   /  /
Description .. : Edition des réductions
Mots clefs ... : PAIE;REDUCTION
*****************************************************************
PT1   : 16/01/2002 SB V571 Fiche de bug n°414 : Requête re-interpretée par la
                           fiche
PT2   : 18/04/2002 SB V571 Fiche de bug n°10087 : V_PGI_env.LibDossier non
                           renseigné en Mono
PT3   : 20/11/2002 SB V591 Traitements Etats chainés
PT4   : 11/09/2003 SB V_42 FQ 10811 Intégration de l'état Loi Fillon
PT5-1 : 04/12/2003 SB V_50 FQ 10244 Ajout Rupture périodique
PT5-2 : 04/12/2003 SB V_50 FQ 10982 Suppression option monnaie inversée
PT6   : 28/04/2004 SB V_50 FQ 10244
PT7   : 10/06/2004 VG V_50 Nouvel état Réduction loi Fillon Notaire CRPCEN
                           FQ 11338
PT8   : 22/11/2005 SB V_650 FQ 12688 ajout des critères salariés
PT9   : 19/05/2006 PH V_650 FQ 12688 Suite Ajout critères zones libres en multicombo
PT10  : 05/12/2007 FC V_810 FQ 14940 + passage par une tob
PT11  : 11/12/2007 FC V_810 FQ 15047 Prendre le cumul 44 par défaut et non le cumul 42
PT12  : 17/06/2008 FC V_810 FQ 15369 Ajouter une colonne "rémunérations" sur l'état "heures sup exo"
PT15  : 02/10/2008 FC FQ 15369 Tenir compte du cumul figurant dans les paramsoc 
}
unit UTOFPGEditReduction;

interface
uses StdCtrls, Controls, Classes, sysutils, ComCtrls,
{$IFDEF EAGLCLIENT}
  eQRS1, UtileAGL,
{$ELSE}
{$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}  QRS1,EdtREtat,
{$ENDIF}
  HCtrls, HEnt1, HMsgBox, UTOF, ParamDat,
  ParamSoc, HQry, UTob, UTOFPGEtats;

type
  TOF_PGREDUCTION_ETAT = class(TOF_PGEtats)
  public
    procedure OnArgument(Arguments: string); override;
    procedure OnUpdate; override; //PT1
    procedure OnLoad; override;
  private
    Origine: string;
    DD, DF: TDateTime;
    procedure DateElipsisclick(Sender: TObject);
    procedure Change(Sender: TObject);
    //     procedure MonnaieInverse(Sender: TObject); PT5-2
    procedure AffectRup(Sender: TObject); //PT5-1
    procedure ExitEdit(Sender: TObject); { PT8 }
  end;

implementation

uses PGEditOutils2, P5Def, EntPaie, PgOutils2;

{ TOF_PGREDSALAIRE_ETAT }

procedure TOF_PGREDUCTION_ETAT.OnArgument(Arguments: string);
var
  Check: TCheckBox;
  CDd, CDf, Defaut: THEdit;
  Ok: Boolean;
  Min, Max, DebPer, FinPer, ExerPerEncours: string;
  Pref, Ch1, Ch2, Ch3, Cs1, Cs2, Cs3: string;
  Num: Integer;
begin
  inherited;
  //DEB PT3
  if TFQRS1(Ecran).FCodeEtat = '' then TFQRS1(Ecran).FCodeEtat := 'PBS';
  TFQRS1(Ecran).FEtat.Value := TFQRS1(Ecran).FCodeEtat;
  TFQRS1(Ecran).FEtatClick(nil);
  DF := idate1900;
  if (Arguments <> '') then
  begin
    if (Pos('CHAINES', Arguments) > 0) then
    begin
      Origine := ReadTokenSt(Arguments);
      if trim(Arguments) <> '' then
      begin
        DD := StrToDate(ReadTokenSt(Arguments));
        DF := StrToDate(ReadTokenSt(Arguments));
      end;
    end;
  end;
  //FIN PT3
  if TFQRS1(Ecran).FEtat <> nil then
    TFQRS1(Ecran).FEtat.OnChange := Change;

  //Valeur par défaut
  Defaut := ThEdit(getcontrol('DOSSIER'));
  if Defaut <> nil then
  begin
    //   Defaut.text:=V_PGI_env.LibDossier;    //PT2 Mise en commmentaire
    Defaut.text := GetParamSoc('SO_LIBELLE');
    Defaut.visible := False;
  end;
  RecupMinMaxTablette('PG', 'ETABLISS', 'ET_ETABLISSEMENT', Min, Max);
  Defaut := ThEdit(getcontrol('PHB_ETABLISSEMENT'));
  if Defaut <> nil then Defaut.text := Min;
  Defaut := ThEdit(getcontrol('PHB_ETABLISSEMENT_'));
  if Defaut <> nil then Defaut.text := Max;
  CDd := THEdit(GetControl('XX_VARIABLEDEB')); //DEB PT1 Modification des noms de champs
  CDf := THEdit(GetControl('XX_VARIABLEFIN')); //FIN PT1
  if CDd <> nil then CDd.OnElipsisClick := DateElipsisclick;
  if CDf <> nil then CDf.OnElipsisClick := DateElipsisclick;
  ok := RendPeriodeEnCours(ExerPerEncours, DebPer, FinPer);
  if Ok = True then
  begin
    if CDd <> nil then CDd.text := DebPer;
    if CDf <> nil then CDf.text := FinPer;
  end;

  {PT5-2 Mise en commentaire
  Convert:=ThEdit(getcontrol('XX_VARIABLECONV'));
  if convert<>nil then Convert.text:='1';
  ChMonInv := TCheckBox(GetControl('CHMONNAIEINV'));
  if ChMonInv<>nil then ChMonInv.OnClick:=MonnaieInverse;}

  Check := TCheckBox(GetControl('CKEURO'));
  if Check <> nil then Check.Checked := VH_Paie.PGTenueEuro;

  { DEB PT5-1 }
  Check := TCheckBox(GetControl('CKPERIODE'));
  if Check <> nil then Check.OnClick := AffectRup;
  SetControlText('XX_RUPTURE1', DateToStr(idate1900));
  { FIN PT5-1 }


  //DEB PT3 Affect critère standard
  if (origine = 'CHAINES') and (DF <> idate1900) then
  begin
    if CDd <> nil then CDd.text := DateToStr(DD);
    if CDf <> nil then CDf.text := DateToStr(DF);
  end;
  //FIN PT3
  { DEB PT8 }
  RecupMinMaxTablette('PG', 'SALARIES', 'PSA_SALARIE', Min, Max);
  Defaut := ThEdit(getcontrol('PHB_SALARIE'));
  if Defaut <> nil then begin Defaut.text := Min; Defaut.OnExit := ExitEdit; end;
  Defaut := ThEdit(getcontrol('PHB_SALARIE_'));
  if Defaut <> nil then begin Defaut.text := Max; Defaut.OnExit := ExitEdit; end;
  { FIN PT8 }
  // DEB PT9
  SetControlProperty('TBCOMPLEMENT', 'Tabvisible', (VH_Paie.PGNbreStatOrg > 0) or (VH_Paie.PGLibCodeStat <> ''));
  Pref := 'PHB';
  Cs1 := Pref + '_CODESTAT';
  Cs2 := 'T' + Pref + '_CODESTAT';
  Cs3 := 'R_CODESTAT';
  VisibiliteStat(GetControl(Cs1), GetControl(Cs2), GetControl(Cs3));
  VisibiliteStat(GetControl(Cs1 + '_'), GetControl(Cs2 + '_'));
  VisibiliteStat(GetControl(Cs1 + '__'), GetControl(Cs2 + '__'));
  for Num := 1 to 4 do
  begin
    if Num > 4 then Break;
    Ch1 := Pref + '_TRAVAILN' + IntToStr(Num);
    Ch2 := 'T' + Pref + '_TRAVAILN' + IntToStr(Num);
    if (Ch1 <> '') and (Ch2 <> '') then
    begin
      Ch3 := 'R_TRAVAILN' + IntToStr(Num);
      VisibiliteChampSalarie(IntToStr(Num), GetControl(Ch1), GetControl(Ch2), GetControl(Ch3));
      VisibiliteChampSalarie(IntToStr(Num), GetControl(Ch1 + '_'), GetControl(Ch2 + '_'));
      VisibiliteChampSalarie(IntToStr(Num), GetControl(Ch1 + '__'), GetControl(Ch2 + '__'));
    end;
  end;
  for Num := 1 to 4 do
  begin
    if Num > 4 then Break;
    Ch1 := Pref + '_LIBREPCMB' + IntToStr(Num);
    Ch2 := 'T' + Pref + '_LIBREPCMB' + IntToStr(Num);
    if (Ch1 <> '') and (Ch2 <> '') then
    begin
      Ch3 := 'R_LIBREPCMB' + IntToStr(Num);
      VisibiliteChampLibreSal(IntToStr(Num), GetControl(Ch1), GetControl(Ch2), GetControl(Ch3));
      VisibiliteChampLibreSal(IntToStr(Num), GetControl(Ch1 + '_'), GetControl(Ch2 + '_'));
      VisibiliteChampLibreSal(IntToStr(Num), GetControl(Ch1 + '__'), GetControl(Ch2 + '__'));
    end;
  end;
  // FIN PT9
end;

procedure TOF_PGREDUCTION_ETAT.DateElipsisclick(Sender: TObject);
var
  key: char;
begin
  key := '*';
  ParamDate(Ecran, Sender, Key);
end;

procedure TOF_PGREDUCTION_ETAT.Change(Sender: TObject);
begin
  if TFQRS1(Ecran).FEtat <> nil then
    TFQRS1(Ecran).Caption := TFQRS1(Ecran).FEtat.Text;
  UpdateCaption(TFQRS1(Ecran));
end;

{PT5-2 Mise en commentaire
procedure TOF_PGREDUCTION_ETAT.MonnaieInverse(Sender: TObject);
var
ChMonInv : TCheckBox;
Convert:ThEdit;
begin
ChMonInv := TCheckBox(GetControl('CHMONNAIEINV'));
Convert:=ThEdit(getcontrol('XX_VARIABLECONV'));
if (ChMonInv<>nil) and (convert<>nil)  then
  begin
  if ChMonInv.checked=False then Convert.text:='1';
  if ChMonInv.checked=True  then Convert.text:=RendTauxConvertion;
  End;
end;}

//DEB PT1

procedure TOF_PGREDUCTION_ETAT.OnUpdate;
var
  SQL, Temp, Tempo, Critere, AjoutChamp, AjoutDate: string;
  Pages: TPageControl;
  DateDeb, DateFin: THEdit;
  x: integer;
  TobEtat,T,TobHistoBul : Tob;
  Montant,Heure,Allegement,Reduction:double;
  Q : TQuery;
  Cumul:String;
  i:integer;
  DtDeb,DtFin:TDateTime;
  Remuneration:double;//PT12
begin
  inherited;
//-------------------------
  DateDeb := THEdit(GetControl('XX_VARIABLEDEB'));
  DateFin := THEdit(GetControl('XX_VARIABLEFIN'));
  Pages := TPageControl(GetControl('Pages'));
  Temp := RecupWhereCritere(Pages);
  tempo := '';
  critere := '';
  x := Pos('(', Temp);
  if x > 0 then
    Tempo := copy(Temp, x, (Length(temp) - 5));

  if tempo <> '' then
    critere := 'AND ' + Tempo;
{ DEB PT6 }
  if GetControlText('CKPERIODE') = 'X' then
    AjoutDate := 'YEAR(PHB_DATEFIN) ANNEEFIN, MONTH(PHB_DATEFIN) MOISFIN,'
  else
    AjoutDate := ''; //PT5-1

  if GetControlText('CKPERIODE') = 'X' then
    AjoutChamp := 'YEAR(PHB_DATEFIN), MONTH(PHB_DATEFIN),'
  else
    AjoutChamp := ''; //PT5-1

  if (IsValidDate(DateDeb.text)) and (IsValidDate(DateFin.text)) then
{ FIN PT6 }
  begin
    //DEB PT10
    if (TFQRS1(Ecran).FCodeEtat = 'PHS') then
    begin
      if GetControlText('CKPERIODE') = 'X' then
      begin
        AjoutDate := 'PHB_DATEDEBUT,PHB_DATEFIN,'+ AjoutDate;
        AjoutChamp := 'PHB_DATEDEBUT,PHB_DATEFIN,'+ AjoutChamp;
      end;
      SQL := 'SELECT DISTINCT ' + AjoutDate + ' PHB_ETABLISSEMENT,PHB_SALARIE,ETB_HORAIREETABL' +
        ' FROM HISTOBULLETIN' +
        ' LEFT JOIN ETABCOMPL ON PHB_ETABLISSEMENT=ETB_ETABLISSEMENT' +
        ' LEFT JOIN COTISATION ON PHB_NATURERUB=PCT_NATURERUB' +
        ' AND ##PCT_PREDEFINI## PHB_RUBRIQUE=PCT_RUBRIQUE ';
      if (Temp <> '') then
        SQL := SQL + Temp + ' AND '
      else
        SQL := SQL + ' WHERE ';
      SQL := SQL + ' (PHB_MTPATRONAL<>0 OR PHB_RUBRIQUE = "9900")' +
        ' AND (PCT_ALLEGEMENTA2="X" OR PCT_MAJORATA2="X" OR PCT_MINORATA2="X")'+
        ' AND PHB_DATEFIN>="' + UsDateTime(StrToDate(DateDeb.text)) + '"'+
        ' AND PHB_DATEFIN<="' + UsDateTime(StrToDate(DateFin.text)) + '"' +
        ' GROUP BY ' + AjoutChamp + ' PHB_ETABLISSEMENT, PHB_SALARIE, ETB_HORAIREETABL' +
        ' ORDER BY ' + AjoutChamp + ' PHB_ETABLISSEMENT, PHB_SALARIE';
      TobHistoBul := Tob.Create('Leshistobul',Nil,-1);
      TobHistoBul.LoadDetailDBFromSQL ('Leshistobul', SQL);
      if TobEtat <> nil then FreeAndNil(TobEtat);
      TobEtat := Tob.Create('Lessalariespouredition', nil, -1);
      For i := 0 to TobHistoBul.Detail.Count - 1 do
      begin
        T := Tob.Create('Lesfillessalaries', TobEtat, -1);
        T.AddChampSup('PERIODE', False);
        T.AddChampSup('MOISFIN', False);
        T.AddChampSup('ANNEEFIN', False);
        T.AddChampSup('PHB_ETABLISSEMENT', False);
        T.AddChampSup('PHB_SALARIE', False);
        T.AddChampSup('ETB_HORAIREETABL', False);
        T.AddChampSup('HEURES', False);
        T.AddChampSup('MONTANT', False);
        T.AddChampSup('ALLEGEMENT', False);
        T.AddChampSup('REDUCTION', False);
        T.AddChampSup('REMUNERAT', False);  //PT12

        if GetControlText('CKPERIODE') = 'X' then
        begin
          T.PutValue('PERIODE', FloatToStr(TobHistoBul.Detail[i].GetValue('ANNEEFIN')) + ' ' + FloatToStr(TobHistoBul.Detail[i].GetValue('MOISFIN')));
          T.PutValue('MOISFIN', TobHistoBul.Detail[i].GetValue('MOISFIN'));
          T.PutValue('ANNEEFIN', TobHistoBul.Detail[i].GetValue('ANNEEFIN'));
        end
        else
        begin
          T.PutValue('PERIODE', '');
          T.PutValue('MOISFIN', '');
          T.PutValue('ANNEEFIN', '');
        end;
        T.PutValue('PHB_ETABLISSEMENT', TobHistoBul.Detail[i].GetValue('PHB_ETABLISSEMENT'));
        T.PutValue('PHB_SALARIE', TobHistoBul.Detail[i].GetValue('PHB_SALARIE'));
        T.PutValue('ETB_HORAIREETABL', TobHistoBul.Detail[i].GetValue('ETB_HORAIREETABL'));

        if GetControlText('CKPERIODE') = 'X' then
        begin
          DtDeb := TobHistoBul.Detail[i].GetValue('PHB_DATEDEBUT');
          DtFin := TobHistoBul.Detail[i].GetValue('PHB_DATEFIN');
        end
        else
        begin
          DtDeb := StrToDate(DateDeb.text);
          DtFin := StrToDate(DateFin.text);
        end;

        Montant := 0;
        Heure := 0;
        Allegement := 0;
        Reduction := 0;
        Remuneration := 0; //PT12

        //Montant
        SQL := 'SELECT SUM(PHB_BASECOT) AS MONTANT FROM HISTOBULLETIN ' +
          ' LEFT JOIN COTISATION ON PHB_NATURERUB=PCT_NATURERUB ' +
          ' AND ##PCT_PREDEFINI## PHB_RUBRIQUE=PCT_RUBRIQUE ' +
          ' WHERE (PCT_ALLEGEMENTA2="X" AND PHB_DATEFIN>="' + UsDateTime(DtDeb) + '"' +
          ' AND PHB_DATEFIN<="' + UsDateTime(DtFin) + '" AND PHB_MTSALARIAL<>0 ) ' +
          ' AND PHB_SALARIE="' + TobHistoBul.Detail[i].GetValue('PHB_SALARIE') + '"';
        Q := OpenSQL(SQL,True);
        if not Q.Eof then
          Montant := Q.FindField('MONTANT').AsFloat;
         Ferme(Q);

        //DEB PT12
        Q := OpenSql('SELECT SUM(PHC_MONTANT) as MONTANT,PHC_CUMULPAIE '+  //PT15
            'FROM HISTOCUMSAL ' +
            'WHERE PHC_ETABLISSEMENT="' + TobHistoBul.Detail[i].GetValue('PHB_ETABLISSEMENT') + '" AND PHC_REPRISE="-" '+
            'AND PHC_DATEFIN>="' + UsDateTime(DtDeb) + '" AND PHC_DATEFIN<="' + UsDateTime(DtFin) + '" ' +
            'AND PHC_SALARIE IN (SELECT PHB_SALARIE FROM HISTOBULLETIN LEFT JOIN COTISATION '+
            'ON PHB_NATURERUB=PCT_NATURERUB AND ##PCT_PREDEFINI## PHB_RUBRIQUE=PCT_RUBRIQUE '+
            'WHERE PHB_SALARIE="' + TobHistoBul.Detail[i].GetValue('PHB_SALARIE') + '"'+
            ' AND PHB_DATEDEBUT=PHC_DATEDEBUT AND PHB_DATEFIN=PHC_DATEFIN AND PCT_REDUCBASSAL="X" '+
            'AND PHB_DATEFIN>="' + UsDateTime(DtDeb) + '" AND PHB_DATEFIN<="' + UsDateTime(DtFin) + '" ' +
            'AND PHB_SALARIE=PHC_SALARIE AND PHB_ETABLISSEMENT=PHC_ETABLISSEMENT  AND PHB_MTPATRONAL<>0 ) '+
            'GROUP BY PHC_SALARIE,PHC_CUMULPAIE', TRUE);
        while not Q.eof do                                                            //PT15
        begin                                                                         //PT15
          if Q.FindField('PHC_CUMULPAIE').AsString = GetParamSoc('SO_PGREDREM') then  //PT15
            Remuneration := Remuneration + Q.FindField('MONTANT').AsFloat;            //PT15
          Q.Next;                                                                     //PT15
        end;
        Ferme(Q);
        //FIN PT12

        //Heures
        SQL := 'SELECT SUM(PHB_BASECOT) AS MONTANT FROM HISTOBULLETIN ' +
          ' LEFT JOIN COTISATION ON PHB_NATURERUB=PCT_NATURERUB ' +
          ' AND ##PCT_PREDEFINI## PHB_RUBRIQUE=PCT_RUBRIQUE ' +
          ' WHERE ( PCT_MAJORATA2="X" AND PHB_DATEFIN>="' + UsDateTime(DtDeb) + '"' +
          ' AND PHB_DATEFIN<="' + UsDateTime(DtFin) + '" AND PHB_MTPATRONAL<>0 ) ' +
          ' AND PHB_SALARIE="' + TobHistoBul.Detail[i].GetValue('PHB_SALARIE') + '"';
        Q := OpenSQL(SQL,True);
        if not Q.Eof then
          Heure := Q.FindField('MONTANT').AsFloat;
        Ferme(Q);

        if (Heure = 0) and (Montant <> 0) then
        begin
          Cumul:= GetParamSocSecur ('SO_PGCUMULHEURESSUP', '44');   //PT11
          SQL := ' SELECT SUM (PHC_MONTANT) AS TOTAL'+
                 ' FROM HISTOCUMSAL WHERE'+
                 ' PHC_SALARIE="'+TobHistoBul.Detail[i].GetValue('PHB_SALARIE')+'" AND'+
                 ' PHC_CUMULPAIE="'+Cumul+'" AND'+
                 ' PHC_DATEFIN<="'+UsDateTime(DtFin)+'" AND'+
                 ' PHC_DATEDEBUT>="'+UsDateTime(DtDeb)+'"';
          if (ExisteSQL(SQL)) then
          begin
             Q := OpenSQL(SQL,True);
             Heure := Q.FindField('TOTAL').AsFloat;
             Ferme(Q);
          end;
        end;

        //Allègement salarial
        SQL := 'SELECT  SUM(PHB_MTSALARIAL) AS MONTANT FROM HISTOBULLETIN ' +
          ' LEFT JOIN COTISATION ON PHB_NATURERUB=PCT_NATURERUB ' +
          ' AND ##PCT_PREDEFINI## PHB_RUBRIQUE=PCT_RUBRIQUE ' +
          ' WHERE (PCT_ALLEGEMENTA2="X" AND PHB_DATEFIN>="' + UsDateTime(DtDeb) + '" ' +
          ' AND PHB_DATEFIN<="' + UsDateTime(DtFin) + '" AND PHB_MTSALARIAL<>0 ) ' +
          ' AND PHB_SALARIE="' + TobHistoBul.Detail[i].GetValue('PHB_SALARIE') + '"';
        Q := OpenSQL(SQL,True);
        if not Q.Eof then
          Allegement := Q.FindField('MONTANT').AsFloat;
        Ferme(Q);

        //Réduction patronale
        SQL := 'SELECT  SUM(PHB_MTPATRONAL) AS MONTANT FROM HISTOBULLETIN ' +
          ' LEFT JOIN COTISATION ON PHB_NATURERUB=PCT_NATURERUB ' +
          ' AND ##PCT_PREDEFINI## PHB_RUBRIQUE=PCT_RUBRIQUE ' +
          ' WHERE (PCT_MAJORATA2="X" AND PHB_DATEFIN>="' + UsDateTime(DtDeb) + '" ' +
          ' AND PHB_DATEFIN<="' + UsDateTime(DtFin) + '" AND PHB_MTPATRONAL<>0 ) ' +
          ' AND PHB_SALARIE="' + TobHistoBul.Detail[i].GetValue('PHB_SALARIE') + '"';
        Q := OpenSQL(SQL,True);
        if not Q.Eof then
          Reduction := Q.FindField('MONTANT').AsFloat;
        Ferme(Q);

        T.PutValue('HEURES', Heure);
        T.PutValue('MONTANT', Montant);
        T.PutValue('ALLEGEMENT', Allegement);
        T.PutValue('REDUCTION', Reduction);
        T.PutValue('REMUNERAT', Remuneration); //PT12
      end;
      FreeAndNil(TobHistoBul);
      TFQRS1(Ecran).NatureEtat := 'PRD';
      TFQRS1(Ecran).CodeEtat := 'PHS';
      TFQRS1(Ecran).LaTob := TobEtat;
    end
    //FIN PT10
    else
    begin
      if TFQRS1(Ecran).FCodeEtat = 'PBS' then
        SQL := 'SELECT ' + AjoutDate + ' PHB_SALARIE, PHB_ETABLISSEMENT,' + { PT6 }
          ' SUM(PHB_MTPATRONAL) AS MTPAT, ETB_HORAIREETABL' +
          ' FROM HISTOBULLETIN' +
          ' LEFT JOIN ETABCOMPL ON' +
          ' PHB_ETABLISSEMENT=ETB_ETABLISSEMENT' +
          ' LEFT JOIN COTISATION ON' +
          ' PHB_NATURERUB=PCT_NATURERUB AND' +
          ' ##PCT_PREDEFINI## PHB_RUBRIQUE=PCT_RUBRIQUE WHERE' +
          ' PHB_MTPATRONAL<>0 AND' +
          ' PCT_REDUCBASSAL="X" AND' +
          ' PHB_DATEFIN>="' + UsDateTime(StrToDate(DateDeb.text)) + '" AND' +
          ' PHB_DATEFIN<="' + UsDateTime(StrToDate(DateFin.text)) + '" ' + Critere +
          ' GROUP BY ' + AjoutChamp + ' PHB_ETABLISSEMENT, PHB_SALARIE,' +
          ' ETB_HORAIREETABL' +
          ' ORDER BY ' + AjoutChamp + ' PHB_ETABLISSEMENT, PHB_SALARIE'
      else if TFQRS1(Ecran).FCodeEtat = 'PLA' then
        SQL := 'SELECT DISTINCT ' + AjoutDate + ' PHB_ETABLISSEMENT, PHB_SALARIE,' +
          ' ETB_HORAIREETABL' + { PT6 }
          ' FROM HISTOBULLETIN' +
          ' LEFT JOIN ETABCOMPL ON' +
          ' PHB_ETABLISSEMENT=ETB_ETABLISSEMENT' +
          ' LEFT JOIN COTISATION ON' +
          ' PHB_NATURERUB=PCT_NATURERUB AND' +
          ' ##PCT_PREDEFINI## PHB_RUBRIQUE=PCT_RUBRIQUE WHERE' +
          ' PHB_MTPATRONAL<>0 AND' +
          ' (PCT_ALLEGEMENTA2="X" OR PCT_MAJORATA2="X" OR' +
          ' PCT_MINORATA2="X") AND' +
          ' PHB_DATEFIN>="' + UsDateTime(StrToDate(DateDeb.text)) + '" AND' +
          ' PHB_DATEFIN<="' + UsDateTime(StrToDate(DateFin.text)) + '" ' + Critere +
          ' ORDER BY ' + AjoutChamp + ' PHB_ETABLISSEMENT, PHB_SALARIE'
      else if TFQRS1(Ecran).FCodeEtat = 'PRP' then
        SQL := 'SELECT DISTINCT ' + AjoutDate + ' PHB_ETABLISSEMENT, PHB_SALARIE,' +
          ' ETB_HORAIREETABL, SUM(PHB_MTPATRONAL) AS MTPAT' + { PT6 }
          ' FROM HISTOBULLETIN' +
          ' LEFT JOIN ETABCOMPL ON' +
          ' PHB_ETABLISSEMENT=ETB_ETABLISSEMENT' +
          ' LEFT JOIN COTISATION ON' +
          ' PHB_NATURERUB=PCT_NATURERUB AND' +
          ' ##PCT_PREDEFINI## PHB_RUBRIQUE=PCT_RUBRIQUE WHERE' +
          ' PHB_MTPATRONAL<>0 AND' +
          ' PCT_REDUCREPAS="X" AND' +
          ' PHB_DATEFIN>="' + UsDateTime(StrToDate(DateDeb.text)) + '" AND' +
          ' PHB_DATEFIN<="' + UsDateTime(StrToDate(DateFin.text)) + '" ' + Critere +
          ' GROUP BY ' + AjoutChamp + ' PHB_ETABLISSEMENT, PHB_SALARIE,' +
          ' ETB_HORAIREETABL' +
          ' ORDER BY ' + AjoutChamp + ' PHB_ETABLISSEMENT, PHB_SALARIE'
      else if TFQRS1(Ecran).FCodeEtat = 'PLF' then { DEB PT4 }
        SQL := 'SELECT ' + AjoutDate + ' PHB_SALARIE, ETB_HORAIREETABL,' +
          ' PHB_ETABLISSEMENT, SUM(PHB_MTPATRONAL) AS MTPAT,' + { PT6 }
          ' SUM(PHB_TAUXPATRONAL)/-100 AS COEFFPAT' +
          ' FROM HISTOBULLETIN' +
          ' LEFT JOIN ETABCOMPL ON' +
          ' PHB_ETABLISSEMENT=ETB_ETABLISSEMENT' +
          ' LEFT JOIN COTISATION ON' +
          ' PHB_NATURERUB=PCT_NATURERUB AND' +
          ' ##PCT_PREDEFINI## PHB_RUBRIQUE=PCT_RUBRIQUE WHERE' +
          ' PHB_MTPATRONAL<>0 AND' +
          ' PCT_REDUCBASSAL="X" AND' +
          ' PHB_DATEFIN>="' + UsDateTime(StrToDate(DateDeb.text)) + '" AND' +
          ' PHB_DATEFIN<="' + UsDateTime(StrToDate(DateFin.text)) + '" ' + Critere +
          ' GROUP BY ' + AjoutChamp + ' PHB_ETABLISSEMENT, PHB_SALARIE,' +
          ' ETB_HORAIREETABL' +
          ' ORDER BY ' + AjoutChamp + ' PHB_ETABLISSEMENT, PHB_SALARIE' {FIN PT4}
//PT7
      else if TFQRS1(Ecran).FCodeEtat = 'PFN' then
        SQL := 'SELECT ' + AjoutDate + ' PHB_SALARIE, ETB_HORAIREETABL,' +
          ' PHB_ETABLISSEMENT, SUM(PHB_MTPATRONAL) AS MTPAT,' +
          ' SUM(PHB_TAUXPATRONAL)/-100 AS COEFFPAT ' +
          ' FROM HISTOBULLETIN ' +
          ' LEFT JOIN ETABCOMPL ON' +
          ' PHB_ETABLISSEMENT=ETB_ETABLISSEMENT' +
          ' LEFT JOIN COTISATION ON' +
          ' PHB_NATURERUB=PCT_NATURERUB AND' +
          ' ##PCT_PREDEFINI## PHB_RUBRIQUE=PCT_RUBRIQUE WHERE' +
          ' PHB_MTPATRONAL<>0 AND' +
          ' (PCT_RUBRIQUE="4436" OR PCT_RUBRIQUE="4440") AND' +
          ' PHB_DATEFIN>="' + UsDateTime(StrToDate(DateDeb.text)) + '" AND' +
          ' PHB_DATEFIN<="' + UsDateTime(StrToDate(DateFin.text)) + '" ' + Critere +
          ' GROUP BY ' + AjoutChamp + ' PHB_ETABLISSEMENT, PHB_SALARIE,' +
          ' ETB_HORAIREETABL' +
          ' ORDER BY ' + AjoutChamp + ' PHB_ETABLISSEMENT, PHB_SALARIE';
//FIN PT7
      TFQRS1(Ecran).WhereSQL := SQL;
    end;
  end
  else
    PgiBox('Vous devez saisir une période de debut et de fin de paie.', Ecran.caption);
end;
//FIN PT1

procedure TOF_PGREDUCTION_ETAT.OnLoad;
begin
  inherited;
  TFQRS1(Ecran).FEtatClick(nil);
  //DEB PT3 Affect critère standard
  if (origine = 'CHAINES') and (DF <> idate1900) then
  begin
    SetControlText('XX_VARIABLEDEB', DateToStr(DD));
    SetControlText('XX_VARIABLEFIN', DateToStr(DF));
  end;
  //FIN PT3
end;
{ DEB PT5-1 }

procedure TOF_PGREDUCTION_ETAT.AffectRup(Sender: TObject);
begin
  SetControltext('XX_RUPTURE1', DateToStr(IDate1900));
  if GetControlText('CKPERIODE') = 'X' then
    SetControltext('XX_RUPTURE1', 'ANNEEFIN,MOISFIN'); { PT6 }
end;
{ FIN PT5-1 }

{ DEB PT8 }

procedure TOF_PGREDUCTION_ETAT.ExitEdit(Sender: TObject);
var edit: thedit;
begin
  edit := THEdit(Sender);
  if edit <> nil then
    if edit.text <> '' then
      if (VH_Paie.PgTypeNumSal = 'NUM') and (length(Edit.text) < 11) and (isnumeric(edit.text)) then
        edit.text := AffectDefautCode(edit, 10);
end;
{ FIN PT8 }


initialization
  registerclasses([TOF_PGREDUCTION_ETAT]);
end.

