{***********UNITE*************************************************
Auteur  ...... : FC
Créé le ...... : 30/07/2007
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : PGMULELTDYNGROUPEE ()
Mots clefs ... : TOF;PGMULELTDYNGROUPEE
*****************************************************************
PT1 FC  03/09/2007 V_80 : FQ 14722 prédéfini CEG dans le niveau préconisé
}
Unit UTOFPGMULELTDYNGROUPEE ;

Interface

Uses StdCtrls, 
     Controls, 
     Classes, 
{$IFNDEF EAGLCLIENT}
     db, 
     {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF} 
     HDB,Mul,Fe_Main,
{$else}
     eMul,MainEAGL,
{$ENDIF}
     forms,
     sysutils,
     ComCtrls,
     HCtrls,
     HEnt1,
     HMsgBox,
     UTOF,
     uTob,
     EntPaie,
     P5Def,
     PGOutils2,
     ParamSoc,
     HQry,
     ed_Tools,
     HStatus;
Type
  TOF_PGMULELTDYNGROUPEE = Class (TOF)
  public
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure ExitEdit(Sender: TObject);
  private
    {$IFNDEF EAGLCLIENT}
    Liste : THDBGrid;
    {$ELSE}
    Liste : THGrid;
    {$ENDIF}
    procedure GrilleDblClick(Sender : TObject);
    procedure OnClickSalarieSortie(Sender: TObject);
    procedure ConstruireTobEltDyn;
    procedure RemplirComboChoix(Sender : TObject);
    procedure AlimentationTob();
  end ;
var PGTobEltDyn,PGTobLesElements : Tob;
    PGDateDebEltDyn,PGDateFinEltDyn : TDateTime;
    PGTypElement : String;

Implementation

procedure TOF_PGMULELTDYNGROUPEE.OnLoad ;
var
  DateArret : TdateTime;
  StDateArret : string;
begin
  if  TCheckBox(GetControl('CKSORTIE'))<>nil then
  begin
    if (GetControlText('CKSORTIE')='X') and (IsValidDate(GetControlText('DATEARRET')))then
    begin
      DateArret:=StrtoDate(GetControlText('DATEARRET'));
      StDateArret:=' AND (PSA_DATESORTIE>="'+UsDateTime(DateArret)+'" OR PSA_DATESORTIE="'+UsdateTime(Idate1900)+'" OR PSA_DATESORTIE IS NULL) ';
      StDateArret:=StDateArret + ' AND PSA_DATEENTREE <="'+UsDateTime(DateArret)+'"';
      SetControlText('XX_WHERE',StDateArret);
    end
    else
       SetControlText('XX_WHERE','');
  end
  else
    StDateArret:='';
end ;

procedure TOF_PGMULELTDYNGROUPEE.OnArgument (S : String ) ;
var
  Check : TCheckBox;
  Defaut: THEdit;
  Num : Integer;
  ValCombo : THValComboBox;
begin
  Inherited ;
  {$IFNDEF EAGLCLIENT}
  Liste := THDBGrid(GetControl('FListe'));
  {$ELSE}
  Liste := THGrid(GetControl('FListe'));
  {$ENDIF}
  If Liste <> Nil Then Liste.OnDblClick := GrilleDblClick ;
  SetControlvisible('DATEARRET',True);
  SetControlvisible('TDATEARRET',True);
  SetControlEnabled('DATEARRET',False);
  SetControlEnabled('TDATEARRET',False);
  Check:=TCheckBox(GetControl('CKSORTIE'));
  if Check=nil then
  Begin
    SetControlVisible('DATEARRET',False);
    SetControlVisible('TDATEARRET',False);
  End
  else
    Check.OnClick:=OnClickSalarieSortie;
  For Num := 1 to VH_Paie.PGNbreStatOrg do
  begin
    if Num >4 then Break;
    VisibiliteChampSalarie (IntToStr(Num),GetControl ('PSA_TRAVAILN'+IntToStr(Num)),GetControl ('TPSA_TRAVAILN'+IntToStr(Num)));
  end;
  VisibiliteStat (GetControl ('PSA_CODESTAT'),GetControl ('TPSA_CODESTAT')) ;
  Defaut:=ThEdit(getcontrol('PSA_SALARIE'));
  If Defaut<>nil then Defaut.OnExit:=ExitEdit;
  ValCombo := THValComboBox(getcontrol('TYPEELT'));
  if ValCombo <> nil then ValCombo.OnChange := RemplirComboChoix;
  if not GetParamSocSecur('SO_PGGESTELTDYNDOS',False) then
    SetControlProperty('TYPEELT','Plus',' AND CO_CODE <> "ELN"');
  SetControlText('TYPEELT','ELD');
  PGTypElement := GetControlText('TYPEELT');
end ;

procedure TOF_PGMULELTDYNGROUPEE.ExitEdit(Sender: TObject);
var edit : thedit;
begin
  edit:=THEdit(Sender);
  if edit <> nil then
    if (VH_Paie.PgTypeNumSal='NUM') and (length(Edit.text)<11) and (isnumeric(edit.text)) then
      edit.text:=AffectDefautCode(edit,10);
end;

procedure TOF_PGMULELTDYNGROUPEE.GrilleDblClick(Sender: TObject);
var
  MultiC : THMultivALCombobox;
begin
   MultiC := THMultiValCombobox(GetControl('LISTEELT'));
   If MultiC.TEXT = '' then
   begin
     PGIBox('Vous devez choisir au moins un élément à modifier',Ecran.Caption);
     Exit;
   end;
   if (Liste.NbSelected = 0) and (TFMul(Ecran).BSelectAll.Down = False) then
   begin
     PGIBox('Aucun élément sélectionné', Ecran.Caption);
     Exit;
   end;
   PGTobEltDyn := tob.Create('SaisieEltDyn',Nil,-1);
   ConstruireTobEltDyn;
   AGLLanceFiche('PAY','ELTDYNGROUPEE','','',GetControlText('LISTEELT')+';'+'CREATION');
   Liste.ClearSelected;
   FreeAndNil(PGTobEltDyn);
   FreeAndNil(PGTobLesElements);
end;

procedure TOF_PGMULELTDYNGROUPEE.OnClickSalarieSortie(Sender: TObject);
begin
  SetControlenabled('DATEARRET',(GetControltext('CKSORTIE')='X'));
  SetControlenabled('TDATEARRET',(GetControltext('CKSORTIE')='X'));
end;

procedure TOF_PGMULELTDYNGROUPEE.RemplirComboChoix(Sender : TObject);
var
  StWhere : String;
begin
  SetControlText('LISTEELT','');
  if (GetControlText('TYPEELT') = 'ELD') then
  begin
    SetControlProperty('LISTEELT','Plus','');
    SetControlProperty('LISTEELT','DataType','PGELEMENTDYN');
  end;

  if (GetControlText('TYPEELT') = 'ELN') then
  begin
    SetControlProperty('LISTEELT','DataType','PGELEMENTNAT');
    StWhere := '';
    // Si présence de niveaux alors on en tient compte sinon on prend tous les éléments nationaux
    if ExisteSQL('SELECT PNR_TYPENIVEAU FROM ELTNIVEAUREQUIS WHERE ##PNR_PREDEFINI##') then
//PT1      StWhere := ' AND PEL_CODEELT IN (SELECT DISTINCT PNR_CODEELT FROM ELTNIVEAUREQUIS WHERE ##PNR_PREDEFINI## (PNR_TYPENIVEAU="SAL" OR PNR_NIVMAXPERSO="SAL"))';
      //PT1
      StWhere := ' AND PEL_CODEELT IN (SELECT DISTINCT PNR_CODEELT FROM ELTNIVEAUREQUIS ' +
        ' WHERE ((PNR_PREDEFINI = "CEG" AND (PNR_TYPENIVEAU="SAL" OR PNR_NIVMAXPERSO="SAL")' +
        ' AND PNR_CODEELT NOT IN (SELECT PNR_CODEELT FROM ELTNIVEAUREQUIS WHERE PNR_PREDEFINI="STD"))' +
        ' OR (PNR_PREDEFINI = "STD" AND (PNR_TYPENIVEAU="SAL" OR PNR_NIVMAXPERSO="SAL"))))';

    SetControlProperty('LISTEELT','Plus',StWhere);
  end;
  PGTypElement := GetControlText('TYPEELT');
end;

procedure TOF_PGMULELTDYNGROUPEE.ConstruireTobEltDyn;
var
  MultiC : THMultiValCombobox;
  Q_Mul: THQuery;
  Salarie : String;
  Q : TQuery;
  LeElement,ElementsSelect : String;
  DateDebut,DateFin : TDateTime;
  TH,TLC: Tob;
  c,i : integer;
  NumChamp : Integer;
  PgInfosModif, PgTypeDonne, CodTabl, Libelle, LienAssoc : String;
  CodeElt, ThemeElt, Monetaire, Abrege, RegimeAlsace : String;
  BValeur : Boolean;
begin
  MultiC := THMultiValCombobox(GetControl('LISTEELT'));
  NumChamp := 1;
  DateDebut := StrToDate(GetControlText('DATEDEBUT'));
  DateFin := StrToDate(GetControlText('DATEFIN'));
  PGDateDebEltDyn := DateDebut;
  PGDateFinEltDyn := DateFin;
  ElementsSelect := GetControlText('LISTEELT');
  PGTobLesElements := Tob.Create('LesElements',Nil,-1);
  LeElement := ReadTokenpipe(ElementsSelect,';');
  BValeur := False;
  if not MultiC.Tous then
  begin
    PGTobLesElements.LoadDetailDB('LesElements','','',Nil,False);
    While LeElement <> '' do
    begin
      TLC := Tob.Create('UnElement',PGTobLesElements,-1);
      // Elément dynamique
      if (GetControlText('TYPEELT') = 'ELD') then
      begin
        Q := OpenSQL('SELECT PPP_PGINFOSMODIF, PPP_PGTYPEDONNE, PPP_CODTABL, PPP_LIBELLE, PPP_LIENASSOC FROM PARAMSALARIE ' +
          ' WHERE ##PPP_PREDEFINI## PPP_PGTYPEINFOLS="ZLS" AND PPP_TYPENIVEAU="SAL" AND PPP_PGINFOSMODIF = "' + LeElement + '"',True);
        If Not Q.Eof then
        begin
          PgInfosModif := Q.FindField('PPP_PGINFOSMODIF').AsString;
          PgTypeDonne := Q.FindField('PPP_PGTYPEDONNE').AsString;
          CodTabl := Q.FindField('PPP_CODTABL').AsString;
          Libelle := Q.FindField('PPP_LIBELLE').AsString;
          LienAssoc := Q.FindField('PPP_LIENASSOC').AsString;
        end;
        Ferme(Q);
        TLC.AddChampSupValeur('PGINFOSMODIF',PgInfosModif);
        TLC.AddChampSupValeur('PGTYPEDONNE',PgTypeDonne);
        TLC.AddChampSupValeur('CODTABL',CodTabl);
        TLC.AddChampSupValeur('LIBELLE',Libelle);
        TLC.AddChampSupValeur('LIENASSOC',LienAssoc);
        TLC.AddChampSupValeur('NUMCHAMP',NumChamp);
      end;
      // Elément dossier
      if (GetControlText('TYPEELT') = 'ELN') then
      begin
        Q := OpenSQL('SELECT PEL_CODEELT, PEL_LIBELLE, PEL_THEMEELT, PEL_MONETAIRE, PEL_ABREGE, PEL_REGIMEALSACE FROM ELTNATIONAUX ' +
          ' WHERE ##PEL_PREDEFINI## PEL_CODEELT="' + LeElement + '"',True);
        If Not Q.Eof then
        begin
          CodeElt := Q.FindField('PEL_CODEELT').AsString;
          Libelle := Q.FindField('PEL_LIBELLE').AsString;
          ThemeElt := Q.FindField('PEL_THEMEELT').AsString;
          Monetaire := Q.FindField('PEL_MONETAIRE').AsString;
          Abrege := Q.FindField('PEL_ABREGE').AsString;
          RegimeAlsace := Q.FindField('PEL_REGIMEALSACE').AsString;
        end;                    
        Ferme(Q);
        TLC.AddChampSupValeur('CODEELT',CodeElt);
        TLC.AddChampSupValeur('LIBELLE',Libelle);
        TLC.AddChampSupValeur('THEMEELT',ThemeElt);
        TLC.AddChampSupValeur('MONETAIRE',Monetaire);
        TLC.AddChampSupValeur('ABREGE',Abrege);
        TLC.AddChampSupValeur('REGIMEALSACE',RegimeAlsace);
        TLC.AddChampSupValeur('NUMCHAMP',NumChamp);
      end;
      NumChamp := NumChamp + 1;
      LeElement := ReadTokenpipe(ElementsSelect,';');
    end;
  end
  else
  begin
    // Elément dynamique
    if (GetControlText('TYPEELT') = 'ELD') then
    begin
      Q := OpenSQL('SELECT PPP_PGINFOSMODIF, PPP_PGTYPEDONNE, PPP_CODTABL, PPP_LIBELLE, PPP_LIENASSOC FROM PARAMSALARIE ' +
        ' WHERE ##PPP_PREDEFINI## PPP_PGTYPEINFOLS="ZLS" AND PPP_TYPENIVEAU="SAL"',True);
      PGTobLesElements.LoadDetailDB('LesElements','','',Q,False);
      Ferme(Q);
      for i := 0 to  PGTobLesElements.Detail.Count -1  do
      begin
        TLC := PGTobLesElements.Detail[i];
        PgInfosModif := TLC.GetValue('PPP_PGINFOSMODIF');
        PgTypeDonne := TLC.GetValue('PPP_PGTYPEDONNE');
        CodTabl := TLC.GetValue('PPP_CODTABL');
        Libelle := TLC.GetValue('PPP_LIBELLE');
        LienAssoc := TLC.GetValue('PPP_LIENASSOC');
        TLC.AddChampSupValeur('PGINFOSMODIF',PgInfosModif);
        TLC.AddChampSupValeur('PGTYPEDONNE',PgTypeDonne);
        TLC.AddChampSupValeur('CODTABL',CodTabl);
        TLC.AddChampSupValeur('LIBELLE',Libelle);
        TLC.AddChampSupValeur('LIENASSOC',LienAssoc);
        TLC.AddChampSupValeur('NUMCHAMP',NumChamp);
        NumChamp := NumChamp + 1;
      end;
    end;
    // Elément dossier
    if (GetControlText('TYPEELT') = 'ELN') then
    begin
      if ExisteSQL('SELECT PNR_TYPENIVEAU FROM ELTNIVEAUREQUIS WHERE ##PNR_PREDEFINI##') then
        Q := OpenSQL('SELECT DISTINCT (PEL_CODEELT), PEL_LIBELLE, PEL_THEMEELT, PEL_MONETAIRE, PEL_ABREGE, PEL_REGIMEALSACE FROM ELTNATIONAUX ' +
          ' WHERE ##PEL_PREDEFINI##' +
          ' AND PEL_CODEELT IN (SELECT DISTINCT PNR_CODEELT FROM ELTNIVEAUREQUIS WHERE ##PNR_PREDEFINI## (PNR_TYPENIVEAU="SAL" OR PNR_NIVMAXPERSO="SAL"))' ,True)
      else
        Q := OpenSQL('SELECT DISTINCT (PEL_CODEELT), PEL_LIBELLE, PEL_THEMEELT, PEL_MONETAIRE, PEL_ABREGE, PEL_REGIMEALSACE FROM ELTNATIONAUX ' +
          ' WHERE ##PEL_PREDEFINI##' ,True);
      PGTobLesElements.LoadDetailDB('LesElements','','',Q,False);
      Ferme(Q);
      for i := 0 to  PGTobLesElements.Detail.Count -1  do
      begin
        TLC := PGTobLesElements.Detail[i];
        CodeElt := TLC.GetValue('PEL_CODEELT');
        Libelle := TLC.GetValue('PEL_LIBELLE');
        ThemeElt := TLC.GetValue('PEL_THEMEELT');
        Monetaire := TLC.GetValue('PEL_MONETAIRE');
        Abrege := TLC.GetValue('PEL_ABREGE');
        RegimeAlsace := TLC.GetValue('PEL_REGIMEALSACE');
        TLC.AddChampSupValeur('CODEELT',CodeElt);
        TLC.AddChampSupValeur('LIBELLE',Libelle);
        TLC.AddChampSupValeur('THEMEELT',ThemeElt);
        TLC.AddChampSupValeur('MONETAIRE',Monetaire);
        TLC.AddChampSupValeur('ABREGE',Abrege);
        TLC.AddChampSupValeur('REGIMEALSACE',RegimeAlsace);
        TLC.AddChampSupValeur('NUMCHAMP',NumChamp);
        NumChamp := NumChamp + 1;
      end;
    end;
  end;

  if Liste = nil then Exit;
  Q_Mul := THQuery(Ecran.FindComponent('Q'));
  if Q_Mul = nil then Exit;
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
      AlimentationTob;
      MoveCurProgressForm(Salarie);
    end;
    FiniMoveProgressForm;
  end
  else if liste.AllSelected then
  begin
    {$IFDEF EAGLCLIENT}
    if (TFMul(Ecran).bSelectAll.Down) then TFMul(Ecran).Fetchlestous;
    {$ENDIF}
    InitMoveProgressForm(nil, 'Début du traitement', 'Veuillez patienter SVP ...', TFmul(Ecran).Q.RecordCount, FALSE, TRUE);
    InitMove(TFmul(Ecran).Q.RecordCount, '');
    Q_Mul.First;
    while not Q_Mul.EOF do
    begin
      AlimentationTob;
      Q_Mul.Next;
    end;
    FiniMoveProgressForm;
  end;
end;

procedure TOF_PGMULELTDYNGROUPEE.AlimentationTob;
var
  Salarie,Etab,Popul : String;
  TH : Tob;
  c : integer;
  Q, Qt : TQuery;
  PgInfosModif,NumChamp,PgTypeDonne,CodeElt : String;
  BValeur,BTrouve : Boolean;
  St : String;
  LocalTob,TFind : TOB;
  ConvSal : String;
begin
  Salarie := TFmul(Ecran).Q.FindField('PSA_SALARIE').asstring;
  TH := Tob.Create('LigneElt',PGTobEltDyn,-1);
  TH.AddChampSupValeur('SALARIE',Salarie);
  TH.AddChampSupValeur('LIBELLE',rechDom('PGSALARIE',Salarie,False));
  TH.AddChampSupValeur('DATEAPPLIC',V_PGI.DateEntree);
  // Elément dynamique
  if (GetControlText('TYPEELT') = 'ELD') then
  begin
    For c := 0 to PGTobLesElements.Detail.count - 1 do
    begin
      PgInfosModif := PGTobLesElements.Detail[c].GetValue('PGINFOSMODIF');
      NumChamp := PGTobLesElements.Detail[c].GetValue('NUMCHAMP');
      PgTypeDonne := PGTobLesElements.Detail[c].GetValue('PGTYPEDONNE');

      Q := OpenSQL('SELECT PHD_NEWVALEUR FROM PGHISTODETAIL ' +
        ' WHERE PHD_SALARIE="' + Salarie + '" AND PHD_DATEAPPLIC <= "' + USDATETIME(DEBUTDEMOIS(V_PGI.DateEntree)) + '" AND PHD_PGINFOSMODIF="' + PgInfosModif + '"' +
        ' ORDER BY PHD_DATEAPPLIC DESC',True);
      if (not Q.Eof) then
//      begin
//        If (PgTypeDonne = 'B') then
          TH.AddChampSupValeur(PgInfosModif,Q.FindField('PHD_NEWVALEUR').AsString)
//        else if (PgTypeDonne='S') or (PgTypeDonne='T') then
//          TH.AddChampSupValeur(PgInfosModif,Q.FindField('PHD_NEWVALEUR').AsString)
//        else If PgTypeDonne = 'D' then
//          TH.AddChampSupValeur(PgInfosModif,Q.FindField('PHD_NEWVALEUR').AsString);
//      end
      else
      begin
        BValeur := False;
        If (PgTypeDonne = 'B') then
          TH.AddChampSupValeur(PgInfosModif,BValeur)
        else if (PgTypeDonne='S') or (PgTypeDonne='T') then
          TH.AddChampSupValeur(PgInfosModif,'')
        else If PgTypeDonne = 'D' then
          TH.AddChampSupValeur(PgInfosModif,IDate1900);
      end;
      Ferme(Q);
      TH.AddChampSupValeur('NUMCHAMP'+IntToStr(c+1),NumChamp);
    end;
  end;
  // Elément dossier
  if (GetControlText('TYPEELT') = 'ELN') then
  begin
    For c := 0 to PGTobLesElements.Detail.count - 1 do
    begin
      CodeElt := PGTobLesElements.Detail[c].GetValue('CODEELT');
      NumChamp := PGTobLesElements.Detail[c].GetValue('NUMCHAMP');
      BTrouve := False;
      Etab := TFmul(Ecran).Q.FindField('PSA_ETABLISSEMENT').asstring;

      Popul := '';
{$IFNDEF CPS1}
      Qt := OpenSQL('SELECT PNA_POPULATION FROM SALARIEPOPUL '
        + ' WHERE PNA_SALARIE = "' + Salarie + '"'
        + ' AND PNA_TYPEPOP = "PAI"', True);
      if not Qt.Eof then
        Popul := Qt.FindField('PNA_POPULATION').AsString;
      Ferme(Qt);
{$ENDIF}

      LocalTob := Tob.Create('ELTNATIONDOS', nil, -1);
      St := ' SELECT PED_CODEELT,PED_TYPENIVEAU,PED_MONTANTEURO FROM ELTNATIONDOS ' +
        ' WHERE PED_CODEELT="' + CodeElt + '" AND PED_DATEVALIDITE<="' + USDATETIME(DEBUTDEMOIS(V_PGI.DateEntree)) + '"' +
        ' AND ((PED_TYPENIVEAU="SAL" AND PED_VALEURNIVEAU="' + Salarie + '")' +
        ' OR (PED_TYPENIVEAU="ETB" AND PED_VALEURNIVEAU="' + Etab + '")' +
        ' OR (PED_TYPENIVEAU="POP" AND PED_VALEURNIVEAU="' + Popul + '"))' +
        ' ORDER BY PED_TYPENIVEAU, PED_DATEVALIDITE DESC';
      LocalTob.LoadDetailDBFromSQL('ELTNATIONDOS', St);

      // Recherche si une valeur existe déjà au niveau SAL
      TFind := LocalTob.FindFirst(['PED_CODEELT', 'PED_TYPENIVEAU'], [CodeElt, 'SAL'], false);
      if TFind <> nil then
      begin
        TH.AddChampSupValeur(CodeElt, TFind.GetValue('PED_MONTANTEURO'));
        BTrouve := True;
      end;
      FreeAndNil(TFind);

      // Recherche si une valeur existe déjà au niveau ETB
      if not BTrouve then
      begin
        TFind := LocalTob.FindFirst(['PED_CODEELT', 'PED_TYPENIVEAU'], [CodeElt, 'ETB'], false);
        if TFind <> nil then
        begin
          TH.AddChampSupValeur(CodeElt, TFind.GetValue('PED_MONTANTEURO'));
          BTrouve := True;
        end;
        FreeAndNil(TFind);
      end;

{$IFNDEF CPS1}
      // Recherche si une valeur existe déjà au niveau POP
      if not BTrouve then
      begin
        TFind := LocalTob.FindFirst(['PED_CODEELT', 'PED_TYPENIVEAU'], [CodeElt, 'POP'], false);
        if TFind <> nil then
        begin
          TH.AddChampSupValeur(CodeElt, TFind.GetValue('PED_MONTANTEURO'));
          BTrouve := True;
        end;
        FreeAndNil(TFind);
      end;
{$ENDIF}

      FreeAndNil(LocalTob);

      if not BTrouve then
      begin
        LocalTob := Tob.Create('ELTNATIONAUX', nil, -1);
        St := ' SELECT PEL_CODEELT,PEL_PREDEFINI,PEL_CONVENTION,PEL_MONTANTEURO FROM ELTNATIONAUX ' +
          ' WHERE ##PEL_PREDEFINI## PEL_CODEELT="' + CodeElt + '" AND PEL_DATEVALIDITE<="' + USDATETIME(DEBUTDEMOIS(V_PGI.DateEntree)) + '"' +
          ' ORDER BY PEL_PREDEFINI, PEL_DATEVALIDITE DESC, PEL_CONVENTION';
        LocalTob.LoadDetailDBFromSQL('ELTNATIONAUX', St);

        // Recherche si une valeur existe déjà au niveau DOS
        TFind := LocalTob.FindFirst(['PEL_CODEELT', 'PEL_PREDEFINI'], [CodeElt, 'DOS'], false);
        if TFind <> nil then
        begin
          TH.AddChampSupValeur(CodeElt, TFind.GetValue('PEL_MONTANTEURO'));
          BTrouve := True;
        end;
        FreeAndNil(TFind);

        // Recherche si une valeur existe déjà au niveau STD CNV
        if not BTrouve then
        begin
          ConvSal := '';
          Q := OpenSQL('SELECT PSA_CONVENTION FROM SALARIES WHERE PSA_SALARIE ="'+Salarie+'"',True);
          if (not Q.Eof) then
            ConvSal := Q.FindField('PSA_CONVENTION').AsString;
          Ferme(Q);
          if (ConvSal <> '') then
          begin
            TFind := LocalTob.FindFirst(['PEL_CODEELT', 'PEL_PREDEFINI','PEL_CONVENTION'], [CodeElt, 'STD', ConvSal], false);
            if TFind <> nil then
            begin
              TH.AddChampSupValeur(CodeElt, TFind.GetValue('PEL_MONTANTEURO'));
              BTrouve := True;
            end;
            FreeAndNil(TFind);
          end;
        end;

        // Recherche si une valeur existe déjà au niveau STD 000
        if not BTrouve then
        begin
          TFind := LocalTob.FindFirst(['PEL_CODEELT', 'PEL_PREDEFINI','PEL_CONVENTION'], [CodeElt, 'STD', '000'], false);
          if TFind <> nil then
          begin
            TH.AddChampSupValeur(CodeElt, TFind.GetValue('PEL_MONTANTEURO'));
            BTrouve := True;
          end;
          FreeAndNil(TFind);
        end;

        // Recherche si une valeur existe déjà au niveau CEG
        if not BTrouve then
        begin
          TFind := LocalTob.FindFirst(['PEL_CODEELT', 'PEL_PREDEFINI'], [CodeElt, 'CEG'], false);
          if TFind <> nil then
          begin
            TH.AddChampSupValeur(CodeElt, TFind.GetValue('PEL_MONTANTEURO'));
            BTrouve := True;
          end;
          FreeAndNil(TFind);
        end;
        FreeAndNil(LocalTob);
      end;

      if not BTrouve then
        TH.AddChampSupValeur(CodeElt,0);  //Ne devrait jamais passer là

      TH.AddChampSupValeur('NUMCHAMP'+IntToStr(c+1),NumChamp);
    end;
  end;
end;


Initialization
  registerclasses ( [ TOF_PGMULELTDYNGROUPEE ] ) ;
end.
