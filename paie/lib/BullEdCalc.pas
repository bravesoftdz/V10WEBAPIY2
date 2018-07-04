{
PT1   27/12/2007 FC V_810 FQ 15035 Rajouter la récupération de la date de naissance
                          dans l'édition du bulletin
PT2   02/01/2008 FC V_810 FQ 15044 Pb récupération taux pour cotis de régul
PT3   09/01/2008 FC V_90  FQ 13189 Possibilité que les rubriques non imposables rentrent dans le coût global salarié
PT4   24/01/2008 FC V_81  FQ 15177 Models bulletins personnalisés avant la v 8.03 : plus d'affichage des taux horaires
PT6  17/04/2008 GGU V81 FQ 15361 Gestion uniformisée des zones libres - tables dynamiques
}
unit BullEdCalc;

interface

uses SysUtils
  {$IFDEF VER150}
  ,Variants
  {$ENDIF}
      , HEnt1
      , HCtrls
      , ParamSoc,
  {$IFNDEF EAGLCLIENT}
  {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF}
  {$ENDIF}
      UTOB,StrUtils
      ,ULibEditionPaie
      ;

function CalcOLEEtatBull(sf, sp: string): variant;
function FunctPgPaieEditBulletin(Sf, Sp: string): variant;
function CalcZoneBasTauCoeMnt(ChampPHB,Etab,Salarie,ValChamp,NumZone:String;DateDeb,DateFin:TDateTime):double;
function PGColleZeroDevant(Nombre, Long: integer): string;
procedure FnPaieCreateBull(DateDeb, DateFin: TDateTime; Salarie, Etab, Page, Population: string;DerPage:String);
//function ReadParam(var St: string): string;


implementation

uses
  ULibPGContexte,
  BullEnt,HMsgBox, DB;


function CalcOLEEtatBull(sf, sp: string): variant;
var
  Champ, Val, Auxiliaire, mode, LibelleBanque : String;
  cumul : String;
  QRech: TQuery;
  i : integer;
  DateAncte,Datefin,Datedeb : TDateTime;
  Nb, j, NbA, NbM : integer;
  NbAnneeMois : Double;      
  CMois, CJour, aad, aaf: Word;
  Salarie,Etab : String;
  vAN, vPN, vRN, vBN, vPN1, vAN1, vRN1, vBN1 : double;
  vDDP, vDFP, vDDPN1, vDFPN1 :TDateTime;
  vGblEditBulCp : String;
  ordre, champs, TypeCumulMois: string;
  Nature,rubrique : String;
  Tob_Temp : tob;
  sDate1,sDate2 : String;
  Regul:String; //PT2
  AjoutImpo:String;   //PT3
  iDateDeb,iDateFin :integer;  //PT3
begin
  if (sf = 'PAIECREATEBULL') or (sf = 'CREATEBULL') or (sf = 'PAIEGETVALZONE') then
  begin
    result := FunctPgPaieEditBulletin(sf, sp);
    exit;
  end;

  //DEB PT3
  {@GETMONTANTIMPO() Retourne le montant de la rubrique si on doit l'ajouter dans le coût global}
  if sf = 'GETMONTANTIMPO' then
  begin
    DateDeb := StrToDate(ReadTokenSt(sp));
    DateFin := StrToDate(ReadTokenSt(sp));
    Salarie := Trim(READTOKENST(sp));
    if (isnumeric(Salarie) and (GetParamSocSecur('SO_PGTYPENUMSAL','NUM',false) = 'NUM')) then
      Salarie := PGColleZeroDevant(StrToInt(Salarie), 10);
    Etab := READTOKENST(sp);

    QRech := OpenSql('SELECT SUM(PHB_MTREM) AS MTREM FROM HISTOBULLETIN ' +
      ' WHERE PHB_NATURERUB="AAA"  ' +
      ' AND PHB_SALARIE="' + Salarie + '" AND PHB_ETABLISSEMENT="' + Etab + '" ' +
      ' AND PHB_RUBRIQUE IN (SELECT PRM_RUBRIQUE FROM REMUNERATION WHERE PRM_AJOUTIMPO="X") ' +
      ' AND PHB_DATEDEBUT>="' + UsDateTime(Datedeb) + '" ' +
      ' AND PHB_DATEFIN<="' + UsDateTime(Datefin) + '" ', TRUE);
    if not QRech.eof then //PORTAGECWAS
      result := FloatToStr(QRech.FindField('MTREM').AsFloat)
    else
      Result := 0;
    Ferme(QRech);
  end;
  //FIN PT3

  {@GETCUMULEX([CRI_XX_VARIABLECUM])
  Bulletin de paie , Retourne la somme du Cumul du debut de l'exercice social
  jusqu'à la fin de session de paie éditée
  }
  if sf = 'GETCUMULEX' then
  begin
    cumul := Trim(READTOKENST(sp));
    if cumul <> '' then
    begin
      for i := 1 to TPGContexte.GetCurrent.MaxCum do
      begin
        if TPGContexte.GetCurrent.TabCumul[i] = Cumul then
        begin
          if TPGContexte.GetCurrent.TabCumulEx[i] = 0 then
            result := ''
          else result := TPGContexte.GetCurrent.TabCumulEx[i];
          Break;
        end;
      end;
    end
    else
      result := '';
    exit;
  end;

  //@GETCUM(CUMULPAIE;[PHB_DATEDEBUT];[PHB_DATEFIN];[PHB_SALARIE];[PHB_ETABLISSEMENT])
  //Bulletin de paie  ,
  //Retourne le montant d'un cumul pour une session,un salarié et un établissement
  if sf = 'GETCUM' then
  begin
    result := TPGContexte.GetCurrent.Cumul21;
    exit;
  end;

  //@GETCUMULMOIS([SO_PGCUMULMOIS1])
  if sf = 'GETCUMULMOIS' then
  begin //Soit code cumul soit code remuneration
    ordre := Trim(READTOKENST(sp));
    cumul := Trim(READTOKENST(sp));
    if cumul = '' then
    begin
      result := '';
      exit;
    end;
    if Length(cumul) = 2 then //Cas d'un cumul
    begin
      for i := 1 to TPGContexte.GetCurrent.MaxCum do
      begin
        if TPGContexte.GetCurrent.TabCumul[i] = Cumul then
        begin
          if TPGContexte.GetCurrent.TabCumulMois[i] = 0 then
            result := ''
          else result := TPGContexte.GetCurrent.TabCumulMois[i];
          Break;
        end;
      end;
    end;
    if Length(Cumul) = 4 then //Cas d'une rémunération
    begin
      Champs := '';
      if ordre = '1' then TypeCumulMois := GetParamSocSecur('SO_PGTYPECUMULMOIS1','',false);// VH_Paie.PgTypeCumulMois1;
      if ordre = '2' then TypeCumulMois := GetParamSocSecur('SO_PGTYPECUMULMOIS2','',false);//VH_Paie.PgTypeCumulMois2;
      if ordre = '3' then TypeCumulMois := GetParamSocSecur('SO_PGTYPECUMULMOIS3','',false);//VH_Paie.PgTypeCumulMois3;
      if TypeCumulMois = '05' then Champs := 'PHB_BASEREM';
      if TypeCumulMois = '06' then Champs := 'PHB_TAUXREM';
      if TypeCumulMois = '07' then Champs := 'PHB_COEFFREM';
      if TypeCumulMois = '08' then Champs := 'PHB_MTREM';
      if (Champs <> '') and (Cumul <> '') then
        if (TPGContexte.GetCurrent.CMEtab <> '') and (TPGContexte.GetCurrent.CMSAL <> '') and (TPGContexte.GetCurrent.CMDD > 0) and (TPGContexte.GetCurrent.CMDF > 0) then
        begin
          QRech := OpenSql('SELECT ' + Champs + ' FROM HISTOBULLETIN ' +
            ' WHERE PHB_NATURERUB="AAA"  ' +
            ' AND PHB_SALARIE="' + TPGContexte.GetCurrent.CMSal + '" AND PHB_ETABLISSEMENT="' + TPGContexte.GetCurrent.CMEtab + '" ' +
            ' AND PHB_RUBRIQUE="' + Cumul + '" ' +
            ' AND PHB_DATEDEBUT>="' + UsDateTime(TPGContexte.GetCurrent.CMDD) + '" ' +
            ' AND PHB_DATEFIN<="' + UsDateTime(TPGContexte.GetCurrent.CMDF) + '" ', TRUE);
          if not QRech.eof then //PORTAGECWAS
            result := FloatToStr(QRech.FindField(champs).AsFloat);
          Ferme(QRech);
        end
        else
          result := '';
    end;
    exit;
  end;

  if sf = 'GETORGANISME' then
  begin
    result := TPGContexte.GetCurrent.organisme;
    exit;
  end;

  {@GETRIB(BP;[Aux];[Mode])
  Retourne le rib du salarié en fonction de son code auxiliaire et de son mode de
  reglement
  BP : pour l'utilisation sur l'édition du bulletin de Paie
  }
  if sf = 'GETRIB' then
  begin
    champ := READTOKENST(sp);
    Auxiliaire := READTOKENST(sp);
    mode := READTOKENST(sp);
    LibelleBanque := ''; //PT57 réinitialisation du rib systématique
    if (mode = '008') and (Auxiliaire <> '') and (champ <> '') then
    begin
      if champ <> 'BP' then
      begin
        QRech := OpenSql('SELECT ' + champ + ' FROM RIB WHERE R_AUXILIAIRE="' + Auxiliaire + '" AND R_PRINCIPAL="X"', TRUE);
        if not QRech.eof then Result := QRech.Fields[0].Asstring; //PORTAGECWAS
        Ferme(QRech);
      end
      else
        if champ = 'BP' then
      begin
        LibelleBanque := '';
        val := '';
        QRech := OpenSql('SELECT  R_DOMICILIATION,R_ETABBQ,R_GUICHET,R_NUMEROCOMPTE,R_CLERIB ' +
          'FROM RIB WHERE R_AUXILIAIRE="' + Auxiliaire + '" AND R_SALAIRE="X"', TRUE);
        if not QRech.Eof then //PORTAGECWAS
        begin
          for i := 1 to QRech.FieldCount - 1 do
            val := val + QRech.Fields[i].Asstring + '  ';
          LibelleBanque := QRech.FindField('R_DOMICILIATION').asstring;
        end;
        Result := Trim(Val);
        Ferme(QRech);
      end;
    end
    else
      Result := '';
    TPGContexte.GetCurrent.LibBanque := LibelleBanque;
    exit;
  end;

  //@GETLIBBANQUE(BP;[PPU_AUXILIAIRE];[PPU_PGMODEREGLE])
  if sf = 'GETLIBBANQUE' then
  begin
    result := TPGContexte.GetCurrent.LibBanque;
    exit;
  end;

  // @RECHNIVEAU()   Pour affichage du net imposable // Max ordre bulletin
  if sf = 'RECHNIVEAU' then
  begin
    if TPGContexte.GetCurrent.Niveau = '' then result := '4'
    else result := StrToInt(TPGContexte.GetCurrent.Niveau);
    exit;
  end;

  //Retourne le mutiplicateur  Taux Coeff
  //@GetTauxRem([PHB_NATURERUB],[PHB_RUBRIQUE])
  if sf = 'GETTAUXREM' then
  begin
    result := '';
    if TPGContexte.GetCurrent.Tob_TauxBull = nil then exit;
    Nature := READTOKENST(sp);
    rubrique := READTOKENST(sp);
    Regul := READTOKENST(sp);    //PT2
    if nature <> 'AAA' then exit;
    If Regul = '' then Regul := '...'; //PT4
    Tob_Temp := TPGContexte.GetCurrent.Tob_TauxBull.FindFirst(['PHB_RUBRIQUE','PHB_COTREGUL'], [rubrique,Regul], False); //PT2
    if Tob_Temp <> nil then result := Tob_Temp.Getvalue('MULTIPLICATEUR');
    exit;
  end;

  if sf = 'GETTXTANCIENNETE' then
  begin
    Result := '';
    sDate1 := READTOKENST(sp);
    if sDate1 <> '' then
      DateAncte := StrToDate(sDate1);
    sDate2 := READTOKENST(sp);
    if sDate2 <> '' then
      DateFin   := StrToDate(sDate2);
    if (sDate1 <> '') and (sDate2 <> '') then
    begin
      DiffMoisJour (DateAncte, DateFin, Nb, j);
      NbAnneeMois := Nb/12;
      NbA := StrToInt(FloatToStr(Int(NbAnneeMois)));
      NbM := Round(Frac(NbAnneeMois)*12);
      if (NbA = 0) then
      begin
        if (Nb <> 0) then
          Result := 'Soit ' + IntToStr(Nb) + ' mois';
      end
      else
      begin
        Result := 'Soit ' + IntToStr(NbA);
        if (NbA = 1) then
          Result := Result + ' an'
        else
          Result := Result + ' ans';
        if (NbM <> 0) then
           Result := Result + ' et ' + IntToStr(NbM) + ' mois';
      end;
    end;
  end;

  // Affectation du solde des congés payés periode acquis pris restants
  // @GETN1(DateDeb,DateFin,CodeSal,Etab)
  if sf = 'GETN1' then
  begin
    Datedeb := StrToDate(READTOKENST(sp));
    Datefin := StrToDate(READTOKENST(sp));
    Salarie := Trim(READTOKENST(sp));
    if (isnumeric(Salarie) and (GetParamSocSecur('SO_PGTYPENUMSAL','NUM',false) = 'NUM')) then
      Salarie := PGColleZeroDevant(StrToInt(Salarie), 10);
    Etab := READTOKENST(sp);
    vPN := 0;
    vAN := 0;
    vRN := 0;
    vBN := 0;
    vPN1 := 0;
    vAN1 := 0;
    vRN1 := 0;
    vBN1 := 0;
    vDDP := 0;
    vDFP := 0;
    vDDPN1 := 0;
    vDFPN1 := 0;
    Result := '';
    if GetParamSocSecur('SO_PGCONGES',False) then //  VH_PAie.PGCongesPayes then
    begin
      RechercheCumulCP(Datedeb, Datefin, Salarie, Etab, vPN, vAN, vRN, vBN, vPN1, vAN1, vRN1, vBN1, vDDP, vDFP, vDDPN1, vDFPN1, vGblEditBulCp);
      if (vGblEditBulCp = '') or (vGblEditBulCp = 'N') then
      begin
        TPGContexte.GetCurrent.PN    := vPN;
        TPGContexte.GetCurrent.AN    := vAN;
        TPGContexte.GetCurrent.RN    := vRN;
        TPGContexte.GetCurrent.BN    := vBN;
        TPGContexte.GetCurrent.PN1   := vPN1;
        TPGContexte.GetCurrent.AN1   := vAN1;
        TPGContexte.GetCurrent.RN1   := vRN1;
        TPGContexte.GetCurrent.BN1   := vBN1;
        TPGContexte.GetCurrent.DDP   := vDDP;
        TPGContexte.GetCurrent.DFP   := vDFP;
        TPGContexte.GetCurrent.DDPN1 := vDDPN1;
        TPGContexte.GetCurrent.DFPN1 := vDFPN1;
        TPGContexte.GetCurrent.GblEditBulCp := vGblEditBulCp;
        exit;
      end;
      if (vDDPN1 > idate1900) and (vDFPN1 > iDate1900) then
      begin
        Decodedate(vDDPN1, aad, Cmois, CJour);
        Decodedate(vDFPN1, aaf, CMois, CJour);
        Result := copy(inttoStr(aad), 3, 2) + '/' + Copy(intToStr(aaf), 3, 2);
      end;
    end;
    TPGContexte.GetCurrent.PN    := vPN;
    TPGContexte.GetCurrent.AN    := vAN;
    TPGContexte.GetCurrent.RN    := vRN;
    TPGContexte.GetCurrent.BN    := vBN;
    TPGContexte.GetCurrent.PN1   := vPN1;
    TPGContexte.GetCurrent.AN1   := vAN1;
    TPGContexte.GetCurrent.RN1   := vRN1;
    TPGContexte.GetCurrent.BN1   := vBN1;
    TPGContexte.GetCurrent.DDP   := vDDP;
    TPGContexte.GetCurrent.DFP   := vDFP;
    TPGContexte.GetCurrent.DDPN1 := vDDPN1;
    TPGContexte.GetCurrent.DFPN1 := vDFPN1;
    TPGContexte.GetCurrent.GblEditBulCp := vGblEditBulCp;
    exit;
  end;

  if sf = 'GETN' then
  begin
    Result := '';
    if (TPGContexte.GetCurrent.GblEditBulCp = '') or (TPGContexte.GetCurrent.GblEditBulCp = 'N1') then exit;
    if (TPGContexte.GetCurrent.DDP <> 0) and (TPGContexte.GetCurrent.DFP <> 0) then
    begin
      decodedate(TPGContexte.GetCurrent.DDP, aad, Cmois, CJour);
      Decodedate(TPGContexte.GetCurrent.DFP, aaf, CMois, CJour);
      Result := copy(inttoStr(aad), 3, 2) + '/' + Copy(intToStr(aaf), 3, 2);
    end;
    exit;
  end;

  if sf = 'GETACQN1' then
  begin
    Result := '';
    if (TPGContexte.GetCurrent.GblEditBulCp = '') or (TPGContexte.GetCurrent.GblEditBulCp = 'N') then exit;
    if TPGContexte.GetCurrent.AN1 <> 0 then
      Result := TPGContexte.GetCurrent.AN1;
    exit;
  end;
  if sf = 'GETACQN' then
  begin
    Result := '';
    if (TPGContexte.GetCurrent.GblEditBulCp = '') or (TPGContexte.GetCurrent.GblEditBulCp = 'N1') then exit;
    if TPGContexte.GetCurrent.AN <> 0 then Result := TPGContexte.GetCurrent.AN;
    exit;
  end;
  if sf = 'GETPRN1' then
  begin
    Result := '';
    if (TPGContexte.GetCurrent.GblEditBulCp = '') or (TPGContexte.GetCurrent.GblEditBulCp = 'N') then exit;
    if TPGContexte.GetCurrent.PN1 <> 0 then Result := TPGContexte.GetCurrent.PN1;
    exit;
  end;
  if sf = 'GETPRN' then
  begin
    Result := '';
    if (TPGContexte.GetCurrent.GblEditBulCp = '') or (TPGContexte.GetCurrent.GblEditBulCp = 'N1') then exit;
    if TPGContexte.GetCurrent.PN <> 0 then Result := TPGContexte.GetCurrent.PN;
    exit;
  end;
  if sf = 'GETRESTN1' then
  begin
    Result := '';
    if (TPGContexte.GetCurrent.GblEditBulCp = '') or (TPGContexte.GetCurrent.GblEditBulCp = 'N') then exit;
    if TPGContexte.GetCurrent.RN1 <> 0 then Result := TPGContexte.GetCurrent.RN1;
    exit;
  end;
  if sf = 'GETRESTN' then
  begin
    Result := '';
    if (TPGContexte.GetCurrent.GblEditBulCp = '') or (TPGContexte.GetCurrent.GblEditBulCp = 'N1') then exit;
    if TPGContexte.GetCurrent.RN <> 0 then Result := TPGContexte.GetCurrent.RN;
    exit;
  end;

  //DEB PT1
  if sf = 'GETDATENAISSANCE' then
  begin
    Result := '';
    Salarie := Trim(READTOKENST(sp));
    if (isnumeric(Salarie) and (GetParamSocSecur('SO_PGTYPENUMSAL','NUM',false) = 'NUM')) then
      Salarie := PGColleZeroDevant(StrToInt(Salarie), 10);
    QRech := OpenSql('SELECT PSA_DATENAISSANCE FROM SALARIES WHERE PSA_SALARIE="' + Salarie + '"', TRUE);
    if not QRech.Eof then
      if QRech.FindField('PSA_DATENAISSANCE').AsDateTime <> iDate1900 then
        Result := DateToStr(QRech.FindField('PSA_DATENAISSANCE').AsDateTime);
    Ferme(QRech);
    exit;
  end;
  //FIN PT1

  if sf = 'FREEFN' then
  begin
    Champ := Trim(READTOKENST(sp));
    if Champ = 'PBP' then
    begin
      if TPGContexte.GetCurrent.Tob_TauxBull <> nil then
      begin
        TPGContexte.GetCurrent.Tob_TauxBull.free;
        TPGContexte.GetCurrent.Tob_TauxBull := nil;
      end;
      if sp = 'X' then
      Begin
        If Assigned(TPGContexte.GetCurrent.GblTob_ParamPop) then
        begin
          TPGContexte.GetCurrent.GblTob_ParamPop.Free;
          TPGContexte.GetCurrent.GblTob_ParamPop := nil;
        end;
        If Assigned(TPGContexte.GetCurrent.Tob_ExerciceSocial) then
        begin
          TPGContexte.GetCurrent.Tob_ExerciceSocial.Free;
          TPGContexte.GetCurrent.Tob_ExerciceSocial := nil;
        end;
        If Assigned(TPGContexte.GetCurrent.Tob_DateClotureCp) then
        begin
          TPGContexte.GetCurrent.Tob_DateClotureCp.Free;
          TPGContexte.GetCurrent.Tob_DateClotureCp := nil;
        end;
        If Assigned(TPGContexte.GetCurrent.Tob_OrganismeBull) then
        begin
          TPGContexte.GetCurrent.Tob_OrganismeBull.Free;
          TPGContexte.GetCurrent.Tob_OrganismeBull := nil;
        end;
      end;
    end;
    result := '';
    exit;
  end;

end;

function FunctPgPaieEditBulletin(Sf, Sp: string): variant;
var
  NumZone,Salarie,Etab,Page, Popul :String;
  iDateDeb,iDateFin                 :integer;
  DateDeb,DateFin                   :TDateTime;
  NomChamp              :String;
  Tob_Temp,Tob_Temp2                :Tob;
  Resultat,Resultat1,Resultat2      :double;
  Libelle, Requete,CodeTabLib,St,Lib : String;
  QEltDynamique,Qt,QSalarie,QCumulPaie,Q : TQuery;
  i : integer;
  DerPage:String;
begin
  //Bulletin de paie , pour une session de paie
  //@PAIECREATEBULL([XX_DATEDEBUT];[XX_DATEFIN];[XX_SALARIE];[XX_ETABLISSEMENT];[PAGE])
  if (Sf = 'PAIECREATEBULL') or (Sf = 'CREATEBULL') then
  begin
    iDateDeb := StrToInt(ReadTokenSt(sp));
    iDateFin := StrToInt(ReadTokenSt(sp));
    Salarie := Trim(ReadTokenSt(sp));
    if (isnumeric(Salarie) and (GetParamSocSecur('SO_PGTYPENUMSAL','NUM',false) = 'NUM')) then
      Salarie := PGColleZeroDevant(StrToInt(Salarie), 10);
    Etab := ReadTokenSt(sp);
    Page := ReadTokenSt(sp);
    DerPage := ReadTokenSt(sp);
    Popul := ReadTokenSt(sp);
    TPGContexte.GetCurrent.CMDD := iDatedeb;
    TPGContexte.GetCurrent.CMDF := iDatefin;
    TPGContexte.GetCurrent.CMSal := Salarie;
    TPGContexte.GetCurrent.CMEtab := Etab;
    if (Etab <> '') and (Salarie <> '') and (iDateDeb > 0) and (iDateFin > 0) then
      FnPaieCreateBull(iDateDeb, iDateFin, Salarie, Etab, Page, Popul,DerPage);
    result := '';
    exit;
  end;

  if Sf = 'PAIEGETVALZONE' then
  begin
    // Récupération des paramètres de la fonction @
    NumZone := Trim(ReadTokenSt(sp));
    Salarie := Trim(ReadTokenSt(sp));
    if (isnumeric(Salarie) and (GetParamSocSecur('SO_PGTYPENUMSAL','NUM',false) = 'NUM')) then
      Salarie := PGColleZeroDevant(StrToInt(Salarie), 10);
    Etab := ReadTokenSt(sp);
    if (isnumeric(Etab)) then
      Etab := PGColleZeroDevant(StrToInt(Etab), 3);
    DateDeb := StrToDate(ReadTokenSt(sp));
    DateFin := StrToDate(ReadTokenSt(sp));
    Page := ReadTokenSt(sp);

    Result := '';

    // Inutile de continuer si on a aucun paramètre
    if TPGContexte.GetCurrent.GblTob_ParamPop = nil then exit;

    Resultat  := 0; Resultat1 := 0; Resultat2 := 0;
    if (NumZone <> '') and (Salarie <> '') and (Assigned(TPGContexte.GetCurrent.GblTob_ParamPop)) then
    begin
      NomChamp := 'TYPZONE' + NumZone;
      Tob_Temp := TPGContexte.GetCurrent.GblTob_ParamPop.FindFirst(['PGP_PGNOMCHAMP'], [NomChamp], False);
      if assigned(Tob_Temp) then
      begin
        // -------- ETIQUETTE --------
        if (Tob_Temp.GetValue('PGP_PGVALCHAMP') = 'ETQ') then
        begin
          NomChamp := 'VALZONE' + NumZone;
          Tob_Temp2 := TPGContexte.GetCurrent.GblTob_ParamPop.FindFirst(['PGP_PGNOMCHAMP'], [NomChamp], False);
          if assigned(Tob_Temp2) then
            Result := Tob_Temp2.GetValue('PGP_PGVALCHAMP');
        end;

        // -------- DONNEE FICHE SALARIE --------
        if (Tob_Temp.GetValue('PGP_PGVALCHAMP') = 'SAL') then
        begin
          NomChamp := 'ZSAL' + NumZone;
          Tob_Temp2 := TPGContexte.GetCurrent.GblTob_ParamPop.FindFirst(['PGP_PGNOMCHAMP'], [NomChamp], False);
          if assigned(Tob_Temp2) then
          begin
            //Si on gère l'historisation des données salariés
            result := '';
            if (Tob_Temp2.GetValue('PGP_PGVALCHAMP') = 'COEFFICIENT') then
              Libelle := 'Coefficient';
            if (Tob_Temp2.GetValue('PGP_PGVALCHAMP') = 'QUALIFICATION') then
              Libelle := 'Qualification';
            if (Tob_Temp2.GetValue('PGP_PGVALCHAMP') = 'CODEEMPLOI') then
              Libelle := 'Emploi PCS';
            if (Tob_Temp2.GetValue('PGP_PGVALCHAMP') = 'LIBELLEEMPLOI') then
              Libelle := 'Emploi';
            if (Tob_Temp2.GetValue('PGP_PGVALCHAMP') = 'CONDEMPLOI') then
              Libelle := 'Caractéristique activité';
            if (Tob_Temp2.GetValue('PGP_PGVALCHAMP') = 'DATEANCIENNETE') then
              Libelle := 'Date d''ancienneté';
            if (Tob_Temp2.GetValue('PGP_PGVALCHAMP') = 'DATEENTREE') then
              Libelle := 'Date d''entrée';
            if (Tob_Temp2.GetValue('PGP_PGVALCHAMP') = 'HORAIREMOIS') then
              Libelle := 'Horaire';
            if (Tob_Temp2.GetValue('PGP_PGVALCHAMP') = 'SALAIRETHEO') then
              Libelle := 'Salaire';
            if (Tob_Temp2.GetValue('PGP_PGVALCHAMP') = 'TAUXHORAIRE') then
              Libelle := 'Taux horaire';
            if (Tob_Temp2.GetValue('PGP_PGVALCHAMP') = 'SALAIREMOIS1') then
              Libelle := 'Salaire mensuel';
            if (Tob_Temp2.GetValue('PGP_PGVALCHAMP') = 'SALAIRANN1') then
              Libelle := 'Salaire annuel';
            if (Tob_Temp2.GetValue('PGP_PGVALCHAMP') = 'INDICE') then
              Libelle := 'Indice';
            if (Tob_Temp2.GetValue('PGP_PGVALCHAMP') = 'NIVEAU') then
              Libelle := 'Niveau';
            if (Tob_Temp2.GetValue('PGP_PGVALCHAMP') = 'DADSPROF') then
              Libelle := 'Statut professionnel DADSU';
            if (Tob_Temp2.GetValue('PGP_PGVALCHAMP') = 'DADSCAT') then
              Libelle := 'Catégorie DUCS';
            if (Tob_Temp2.GetValue('PGP_PGVALCHAMP') = 'CATBILAN') then
              Libelle := 'Catégorie bilan social';
            if (Tob_Temp2.GetValue('PGP_PGVALCHAMP') = 'CONVENTION') then
              Libelle := 'Convention collective';
            if (MidStr(Tob_Temp2.GetValue('PGP_PGVALCHAMP'), 1, Length(Tob_Temp2.GetValue('PGP_PGVALCHAMP')) - 1)) = 'TRAVAILN' then
            begin
              if GetParamSocSecur('SO_PGLIBORGSTAT' + (MidStr(Tob_Temp2.GetValue('PGP_PGVALCHAMP'), 9, 1)),'') <> '' then
                Libelle := GetParamSocSecur('SO_PGLIBORGSTAT' + (MidStr(Tob_Temp2.GetValue('PGP_PGVALCHAMP'), 9, 1)),'')
              else
                Libelle := 'Libellé organisation ' + (MidStr(Tob_Temp2.GetValue('PGP_PGVALCHAMP'), 9, 1));
            end;
            if GetParamSocSecur('SO_PGHISTORISATION',True) then // VH_Paie.PgHistorisation then
            begin
              Requete := 'SELECT PHD_NEWVALEUR,PHD_TABLETTE FROM PGHISTODETAIL ' +
                ' WHERE PHD_PGINFOSMODIF = "PSA_' + Tob_Temp2.GetValue('PGP_PGVALCHAMP') + '" ' +
                ' AND PHD_SALARIE = "' + Salarie + '"' +
                ' AND PHD_DATEAPPLIC <= "' + USDATETIME(DateFin) + '"' +
                ' ORDER BY PHD_DATEAPPLIC DESC';
              QEltDynamique := OpenSql(Requete, TRUE);
              if not QEltDynamique.eof then //PORTAGECWAS
              begin
                if QEltDynamique.FindField('PHD_TABLETTE').AsString <> '' then
                begin
                  // Cas particulier pour la qualification car nom tablette tronqué. Il faut la récupérer dans PPP
                  if (Tob_Temp2.GetValue('PGP_PGVALCHAMP') = 'QUALIFICATION') then
                  begin
                    Qt := OpenSql('SELECT PPP_LIENASSOC FROM PARAMSALARIE ' +
                      ' WHERE PPP_PGINFOSMODIF = "PSA_' + Tob_Temp2.GetValue('PGP_PGVALCHAMP') + '"' +
                      ' AND PPP_PREDEFINI = "CEG"', True);
                    if not Qt.Eof then
                      result := Libelle + ' : ' + RechDom(Qt.FindField('PPP_LIENASSOC').AsString,QEltDynamique.FindField('PHD_NEWVALEUR').AsString,False)
                    else
                      result := Libelle + ' : ' + QEltDynamique.FindField('PHD_NEWVALEUR').AsString;
                    Ferme(Qt);
                  end
                  else
                  begin
                    if RechDom(QEltDynamique.FindField('PHD_TABLETTE').AsString,QEltDynamique.FindField('PHD_NEWVALEUR').AsString,False) <> '' then
                      result := Libelle + ' : ' + RechDom(QEltDynamique.FindField('PHD_TABLETTE').AsString,QEltDynamique.FindField('PHD_NEWVALEUR').AsString,False)
                    else
                      result := Libelle + ' : ' + QEltDynamique.FindField('PHD_NEWVALEUR').AsString;
                  end;
                end
                else
                  result := Libelle + ' : ' + QEltDynamique.FindField('PHD_NEWVALEUR').AsString;
              end
              else
                result := '';
              Ferme(QEltDynamique);
            end;
            if (result = '') then
            begin
              Requete := 'SELECT PSA_' + Tob_Temp2.GetValue('PGP_PGVALCHAMP') + ' FROM SALARIES' +
                ' WHERE PSA_SALARIE = "' + Salarie + '"';
              QSalarie := OpenSql(Requete, TRUE);
              if not QSalarie.Eof then //PORTAGECWAS
                if (Tob_Temp2.GetValue('PGP_PGVALCHAMP') = 'TAUXPARTIEL') then
                  Result := Libelle + ' : ' + FloatToStr(QSalarie.Fields[0].AsFloat) + ' %'
                else
                begin
                  if      (Tob_Temp2.GetValue('PGP_PGVALCHAMP') = 'CODEEMPLOI') then
                    Result := Libelle + ' : ' + RechDom('PGCODEPCSESE',QSalarie.Fields[0].AsString,False)
                  else if (Tob_Temp2.GetValue('PGP_PGVALCHAMP') = 'LIBELLEEMPLOI') then
                    Result := Libelle + ' : ' + RechDom('PGLIBEMPLOI',QSalarie.Fields[0].AsString,False)
                  else if (Tob_Temp2.GetValue('PGP_PGVALCHAMP') = 'QUALIFICATION') then
                    Result := Libelle + ' : ' + RechDom('PGLIBQUALIFICATION',QSalarie.Fields[0].AsString,False)
                  else if (Tob_Temp2.GetValue('PGP_PGVALCHAMP') = 'COEFFICIENT') then
                    Result := Libelle + ' : ' + RechDom('PGLIBCOEFFICIENT',QSalarie.Fields[0].AsString,False)
                  else if (Tob_Temp2.GetValue('PGP_PGVALCHAMP') = 'INDICE') then
                    Result := Libelle + ' : ' + RechDom('PGLIBINDICE',QSalarie.Fields[0].AsString,False)
                  else if (Tob_Temp2.GetValue('PGP_PGVALCHAMP') = 'NIVEAU') then
                    Result := Libelle + ' : ' + RechDom('PGLIBNIVEAU',QSalarie.Fields[0].AsString,False)
                  else if (Tob_Temp2.GetValue('PGP_PGVALCHAMP') = 'DADSCAT') then
                    Result := Libelle + ' : ' + RechDom('PGSCATEGORIEL',QSalarie.Fields[0].AsString,False)
                  else if (Tob_Temp2.GetValue('PGP_PGVALCHAMP') = 'DADSPROF') then
                    Result := Libelle + ' : ' + RechDom('PGSPROFESSIONNEL',QSalarie.Fields[0].AsString,False)
                  else if (Tob_Temp2.GetValue('PGP_PGVALCHAMP') = 'CONDEMPLOI') then
                    Result := Libelle + ' : ' + RechDom('PGCONDEMPLOI',QSalarie.Fields[0].AsString,False)
                  else if (Tob_Temp2.GetValue('PGP_PGVALCHAMP') = 'CATBILAN') then
                    Result := Libelle + ' : ' + RechDom('PGCATBILAN',QSalarie.Fields[0].AsString,False)
                  else if (Tob_Temp2.GetValue('PGP_PGVALCHAMP') = 'CONVENTION') then
                    Result := Libelle + ' : ' + RechDom('PGCATBILAN',QSalarie.Fields[0].AsString,False)
                  else if (Tob_Temp2.GetValue('PGP_PGVALCHAMP') = 'TRAVAILN1') then
                    Result := Libelle + ' : ' + RechDom('PGTRAVAILN1',QSalarie.Fields[0].AsString,False)
                  else if (Tob_Temp2.GetValue('PGP_PGVALCHAMP') = 'TRAVAILN2') then
                    Result := Libelle + ' : ' + RechDom('PGTRAVAILN2',QSalarie.Fields[0].AsString,False)
                  else if (Tob_Temp2.GetValue('PGP_PGVALCHAMP') = 'TRAVAILN3') then
                    Result := Libelle + ' : ' + RechDom('PGTRAVAILN3',QSalarie.Fields[0].AsString,False)
                  else if (Tob_Temp2.GetValue('PGP_PGVALCHAMP') = 'TRAVAILN4') then
                    Result := Libelle + ' : ' + RechDom('PGTRAVAILN4',QSalarie.Fields[0].AsString,False)
                  else
                    Result := Libelle + ' : ' + QSalarie.Fields[0].AsString;
                end;
              Ferme(QSalarie);
            end;
          end;
        end;

        if (Tob_Temp.GetValue('PGP_PGVALCHAMP') <> 'ETQ') and (Tob_Temp.GetValue('PGP_PGVALCHAMP') <> 'RES') then
        begin
          NomChamp := 'VALZONE' + NumZone;
          Tob_Temp2 := TPGContexte.GetCurrent.GblTob_ParamPop.FindFirst(['PGP_PGNOMCHAMP'], [NomChamp], False);
          if assigned(Tob_Temp2) then
          begin
            // --------  CUMUL MOIS --------
            if (Tob_Temp.GetValue('PGP_PGVALCHAMP') = 'CMM') then
            begin
              Requete := 'SELECT SUM(PHC_MONTANT) AS MONTANT ' +
                'FROM HISTOCUMSAL WHERE PHC_SALARIE="' + Salarie + '" AND ' +
                'PHC_CUMULPAIE="' + Tob_Temp2.GetValue('PGP_PGVALCHAMP') + '" ' +
                'AND PHC_DATEDEBUT>="' + UsDateTime(DateDeb) + '" ' +
                'AND PHC_DATEFIN<="' + UsDateTime(DateFin) + '" ';
              QCumulPaie := OpenSql(Requete, TRUE);
              if not QCumulPaie.eof then //PORTAGECWAS
                Resultat := QCumulPaie.FindField('MONTANT').AsFloat;
              Ferme(QCumulPaie);

              TPGContexte.GetCurrent.Tab_ZonesCalculees[strtoint(NumZone)] := Resultat;
              Result := RechDom('PGCUMULPAIE',Tob_Temp2.GetValue('PGP_PGVALCHAMP'),False) + ' : ' + FormatFloat('#,##0.00',Resultat);
            end;

            // -------- ELEMENT DYNAMIQUE LIBRE (ex carte orange) --------
            if (Tob_Temp.GetValue('PGP_PGVALCHAMP') = 'ELD') then
            begin
              Requete := 'SELECT PHD_NEWVALEUR,PHD_CODTABL FROM PGHISTODETAIL ' +
                ' WHERE PHD_PGINFOSMODIF = "' + Tob_Temp2.GetValue('PGP_PGVALCHAMP') + '" ' +
                ' AND PHD_SALARIE = "' + Salarie + '"' +
                ' AND PHD_DATEAPPLIC <= "' + USDATETIME(DateFin) + '"' +
                ' ORDER BY PHD_DATEAPPLIC DESC';
              QEltDynamique := OpenSql(Requete, TRUE);
              if not QEltDynamique.eof then //PORTAGECWAS
              begin
                CodeTabLib := QEltDynamique.FindField('PHD_CODTABL').AsString;
                St := ' SELECT PTD_LIBELLECODE '
                    + ' FROM TABLEDIMDET '
                    + ' WHERE ##PTD_PREDEFINI## '
                    + ' AND PTD_DTVALID IN (SELECT MAX(PTE_DTVALID) '   //PT6
                                         + ' FROM TABLEDIMENT '
                                         + ' WHERE ##PTE_PREDEFINI## '
                                         + ' AND PTD_CODTABL=PTE_CODTABL '
                                         + ' AND PTD_PREDEFINI = PTE_PREDEFINI '
                                         + ' AND PTD_NODOSSIER = PTE_NODOSSIER '
                                         + ' AND PTD_NIVSAIS = PTE_NIVSAIS '
                                         + ' AND PTE_DTVALID<="'+UsDateTime(DateFin)+'") '
                    + ' AND PTD_CODTABL="'+CodeTabLib+'" '
                    + ' AND PTD_VALCRIT1="'+QEltDynamique.FindField('PHD_NEWVALEUR').AsString+'"';
                Q := OpenSql(St, TRUE);
                Lib := '';
                if not Q.eof then //PORTAGECWAS
                  Lib:= Q.FindField('PTD_LIBELLECODE').AsString;
                Ferme(Q);
                if Lib = '' then
                  Lib := QEltDynamique.FindField('PHD_NEWVALEUR').AsString;
                result := RechDom('PGZONEHISTOSAL',Tob_Temp2.GetValue('PGP_PGVALCHAMP'),False) + ' : ' + Lib;
              end
              else
                result := '';
              Ferme(QEltDynamique);
            end;

            // -------- ELEMENT DYNAMIQUE MONTANT -------- (valeur particulière d'un élément national pour un salarié)
            if (Tob_Temp.GetValue('PGP_PGVALCHAMP') = 'ELM') then
            begin
              Requete := 'SELECT PED_MONTANTEURO,PED_MONTANT FROM ELTNATIONDOS ' +
                ' WHERE PED_TYPENIVEAU = "SAL" ' +
                ' AND PED_VALEURNIVEAU = "' + Salarie + '"' +
                ' AND PED_DATEVALIDITE <= "' + UsDateTime(DateDeb) + '"' +
                ' AND PED_CODEELT = "' + Tob_Temp2.GetValue('PGP_PGVALCHAMP') + '"' +
                ' ORDER BY PED_DATEVALIDITE DESC';
              QEltDynamique := OpenSql(Requete, TRUE);
              if not QEltDynamique.eof then //PORTAGECWAS
                Resultat := QEltDynamique.FindField('PED_MONTANTEURO').AsFloat;
              if Resultat = 0 then
                Resultat := QEltDynamique.FindField('PED_MONTANT').AsFloat;
              Ferme(QEltDynamique);

              TPGContexte.GetCurrent.Tab_ZonesCalculees[strtoint(NumZone)] := Resultat;
              Result := RechDom('PGELEMENTNAT',Tob_Temp2.GetValue('PGP_PGVALCHAMP'),False) + ' : ' + FormatFloat('#,##0.00',Resultat);
            end;

            // -------- ZONES LIBRES -------
            if (Tob_Temp.GetValue('PGP_PGVALCHAMP') = 'ZLS') then
            begin
              NomChamp := 'ZLS' + NumZone;
              Tob_Temp2 := TPGContexte.GetCurrent.GblTob_ParamPop.FindFirst(['PGP_PGNOMCHAMP'], [NomChamp], False);
              if assigned(Tob_Temp2) then
              begin
                result := '';
                if GetParamSocSecur('SO_PGHISTORISATION',True) then // VH_Paie.PgHistorisation then
                begin
                  if ExisteSQL('SELECT PPP_PGINFOSMODIF FROM PARAMSALARIE WHERE PPP_PGINFOSMODIF = "PSA_' + Tob_Temp2.GetValue('PGP_PGVALCHAMP') + '"' +
                    ' AND PPP_HISTORIQUE="X" AND ##PPP_PREDEFINI##') then
                    begin
                    Requete := 'SELECT PHD_NEWVALEUR,PHD_TABLETTE FROM PGHISTODETAIL ' +
                      ' WHERE PHD_PGINFOSMODIF = "PSA_' + Tob_Temp2.GetValue('PGP_PGVALCHAMP') + '" ' +
                      ' AND PHD_SALARIE = "' + Salarie + '"' +
                      ' AND PHD_DATEAPPLIC <= "' + USDATETIME(DateFin) + '"' +
                      ' ORDER BY PHD_DATEAPPLIC DESC';
                    QEltDynamique := OpenSql(Requete, TRUE);
                    if not QEltDynamique.eof then //PORTAGECWAS
                    begin
                      // Dates libres
                      if (Tob_Temp2.GetValue('PGP_PGVALCHAMP') = 'DATELIBRE1') then
                        result := GetParamsoc('SO_PGLIBDATE1') + ' : ' + DateTimeToStr(QEltDynamique.FindField('PHD_NEWVALEUR').AsDateTime);
                      if (Tob_Temp2.GetValue('PGP_PGVALCHAMP') = 'DATELIBRE2') then
                        result := GetParamsoc('SO_PGLIBDATE2') + ' : ' + DateTimeToStr(QEltDynamique.FindField('PHD_NEWVALEUR').AsDateTime);
                      if (Tob_Temp2.GetValue('PGP_PGVALCHAMP') = 'DATELIBRE3') then
                        result := GetParamsoc('SO_PGLIBDATE3') + ' : ' + DateTimeToStr(QEltDynamique.FindField('PHD_NEWVALEUR').AsDateTime);
                      if (Tob_Temp2.GetValue('PGP_PGVALCHAMP') = 'DATELIBRE4') then
                        result := GetParamsoc('SO_PGLIBDATE4') + ' : ' + DateTimeToStr(QEltDynamique.FindField('PHD_NEWVALEUR').AsDateTime);

                      // Combos libres
                      if (Tob_Temp2.GetValue('PGP_PGVALCHAMP') = 'LIBREPCMB1') then
                        result := GetParamsoc('SO_PGLIBCOMBO1') + ' : ' + RechDom('PGLIBREPCMB1',QEltDynamique.FindField('PHD_NEWVALEUR').AsString,False);
                      if (Tob_Temp2.GetValue('PGP_PGVALCHAMP') = 'LIBREPCMB2') then
                        result := GetParamsoc('SO_PGLIBCOMBO2') + ' : ' + RechDom('PGLIBREPCMB2',QEltDynamique.FindField('PHD_NEWVALEUR').AsString,False);
                      if (Tob_Temp2.GetValue('PGP_PGVALCHAMP') = 'LIBREPCMB3') then
                        result := GetParamsoc('SO_PGLIBCOMBO3') + ' : ' + RechDom('PGLIBREPCMB3',QEltDynamique.FindField('PHD_NEWVALEUR').AsString,False);
                      if (Tob_Temp2.GetValue('PGP_PGVALCHAMP') = 'LIBREPCMB4') then
                        result := GetParamsoc('SO_PGLIBCOMBO4') + ' : ' + RechDom('PGLIBREPCMB4',QEltDynamique.FindField('PHD_NEWVALEUR').AsString,False);

                      // Cases à cocher libres
                      if (Tob_Temp2.GetValue('PGP_PGVALCHAMP') = 'BOOLLIBRE1') then
                      begin
                        if (QEltDynamique.FindField('PHD_NEWVALEUR').AsString = '-') then
                          Result := GetParamsoc('SO_PGLIBCOCHE1') + ' : Non'
                        else
                          Result := GetParamsoc('SO_PGLIBCOCHE1') + ' : Oui';
                      end;
                      if (Tob_Temp2.GetValue('PGP_PGVALCHAMP') = 'BOOLLIBRE2') then
                      begin
                        if (QEltDynamique.FindField('PHD_NEWVALEUR').AsString = '-') then
                          Result := GetParamsoc('SO_PGLIBCOCHE2') + ' : Non'
                        else
                          Result := GetParamsoc('SO_PGLIBCOCHE2') + ' : Oui';
                      end;
                      if (Tob_Temp2.GetValue('PGP_PGVALCHAMP') = 'BOOLLIBRE3') then
                      begin
                        if (QEltDynamique.FindField('PHD_NEWVALEUR').AsString = '-') then
                          Result := GetParamsoc('SO_PGLIBCOCHE3') + ' : Non'
                        else
                          Result := GetParamsoc('SO_PGLIBCOCHE3') + ' : Oui';
                      end;
                      if (Tob_Temp2.GetValue('PGP_PGVALCHAMP') = 'BOOLLIBRE4') then
                      begin
                        if (QEltDynamique.FindField('PHD_NEWVALEUR').AsString = '-') then
                          Result := GetParamsoc('SO_PGLIBCOCHE4') + ' : Non'
                        else
                          Result := GetParamsoc('SO_PGLIBCOCHE4') + ' : Oui';
                      end;
                    end
                    else
                      result := '';
                    Ferme(QEltDynamique);
                  end;
                end;
                if (result = '') then
                begin
                  Requete := 'SELECT PSA_' + Tob_Temp2.GetValue('PGP_PGVALCHAMP') + ' FROM SALARIES' +
                    ' WHERE PSA_SALARIE = "' + Salarie + '"';
                  QSalarie := OpenSql(Requete, TRUE);
                  if not QSalarie.Eof then //PORTAGECWAS
                    if (Tob_Temp2.GetValue('PGP_PGVALCHAMP') = 'DATELIBRE1') then
                    begin
                      if (QSalarie.Fields[0].AsDateTime <> iDate1900) then
                        Result := GetParamsoc('SO_PGLIBDATE1') + ' : ' + DateTimeToStr(QSalarie.Fields[0].AsDateTime);
                    end
                    else if (Tob_Temp2.GetValue('PGP_PGVALCHAMP') = 'DATELIBRE2') then
                    begin
                      if (QSalarie.Fields[0].AsDateTime <> iDate1900) then
                        Result := GetParamsoc('SO_PGLIBDATE2') + ' : ' + DateTimeToStr(QSalarie.Fields[0].AsDateTime);
                    end
                    else if (Tob_Temp2.GetValue('PGP_PGVALCHAMP') = 'DATELIBRE3') then
                    begin
                      if (QSalarie.Fields[0].AsDateTime <> iDate1900) then
                        Result := GetParamsoc('SO_PGLIBDATE3') + ' : ' + DateTimeToStr(QSalarie.Fields[0].AsDateTime);
                    end
                    else if (Tob_Temp2.GetValue('PGP_PGVALCHAMP') = 'DATELIBRE4') then
                    begin
                      if (QSalarie.Fields[0].AsDateTime <> iDate1900) then
                        Result := GetParamsoc('SO_PGLIBDATE4') + ' : ' + DateTimeToStr(QSalarie.Fields[0].AsDateTime);
                    end
                    else if (Tob_Temp2.GetValue('PGP_PGVALCHAMP') = 'BOOLLIBRE1') then
                    begin
                      if (QSalarie.Fields[0].AsString = '-') then
                        Result := GetParamsoc('SO_PGLIBCOCHE1') + ' : Non'
                      else
                        Result := GetParamsoc('SO_PGLIBCOCHE1') + ' : Oui';
                    end
                    else if (Tob_Temp2.GetValue('PGP_PGVALCHAMP') = 'BOOLLIBRE2') then
                    begin
                      if (QSalarie.Fields[0].AsString = '-') then
                        Result := GetParamsoc('SO_PGLIBCOCHE2') + ' : Non'
                      else
                        Result := GetParamsoc('SO_PGLIBCOCHE2') + ' : Oui';
                    end
                    else if (Tob_Temp2.GetValue('PGP_PGVALCHAMP') = 'BOOLLIBRE3') then
                    begin
                      if (QSalarie.Fields[0].AsString = '-') then
                        Result := GetParamsoc('SO_PGLIBCOCHE3') + ' : Non'
                      else
                        Result := GetParamsoc('SO_PGLIBCOCHE3') + ' : Oui';
                    end
                    else if (Tob_Temp2.GetValue('PGP_PGVALCHAMP') = 'BOOLLIBRE4') then
                    begin
                      if (QSalarie.Fields[0].AsString = '-') then
                        Result := GetParamsoc('SO_PGLIBCOCHE4') + ' : Non'
                      else
                        Result := GetParamsoc('SO_PGLIBCOCHE4') + ' : Oui';
                    end
                    else
                    begin
                      if (Tob_Temp2.GetValue('PGP_PGVALCHAMP') = 'LIBREPCMB1') then
                        Result := GetParamsoc('SO_PGLIBCOMBO1') + ' : ' + RechDom('PGLIBREPCMB1',QSalarie.Fields[0].AsString,False)
                      else if (Tob_Temp2.GetValue('PGP_PGVALCHAMP') = 'LIBREPCMB2') then
                        Result := GetParamsoc('SO_PGLIBCOMBO2') + ' : ' + RechDom('PGLIBREPCMB2',QSalarie.Fields[0].AsString,False)
                      else if (Tob_Temp2.GetValue('PGP_PGVALCHAMP') = 'LIBREPCMB3') then
                        Result := GetParamsoc('SO_PGLIBCOMBO3') + ' : ' + RechDom('PGLIBREPCMB3',QSalarie.Fields[0].AsString,False)
                      else if (Tob_Temp2.GetValue('PGP_PGVALCHAMP') = 'LIBREPCMB4') then
                        Result := GetParamsoc('SO_PGLIBCOMBO4') + ' : ' + RechDom('PGLIBREPCMB4',QSalarie.Fields[0].AsString,False)
                      else
                        Result := QSalarie.Fields[0].AsString;
                    end;
                  Ferme(QSalarie);
                end;
              end;
            end;

            // -------- CUMUL PAIE --------
            if (Tob_Temp.GetValue('PGP_PGVALCHAMP') = 'CMP') then
            begin
              for i := 1 to TPGContexte.GetCurrent.MaxCum do
              begin
                if TPGContexte.GetCurrent.TabCumul[i] = Tob_Temp2.GetValue('PGP_PGVALCHAMP') then
                begin
                  Resultat := TPGContexte.GetCurrent.TabCumulEx[i];
                  Break;
                end;
              end;

              TPGContexte.GetCurrent.Tab_ZonesCalculees[strtoint(NumZone)] := Resultat;
              Result := RechDom('PGCUMULPAIE',Tob_Temp2.GetValue('PGP_PGVALCHAMP'),False) + ' : ' + FormatFloat('#,##0.00',Resultat);
            end;

            // -------- BASE, TAUX, COEFFICIENT, MONTANT  DE REMUNERATION --------
            if (Tob_Temp.GetValue('PGP_PGVALCHAMP') = 'BAS') then
            begin
              Result := CalcZoneBasTauCoeMnt('PHB_BASEREM',Etab,Salarie,Tob_Temp2.GetValue('PGP_PGVALCHAMP'),NumZone,DateDeb,DateFin);
              Result := RechDom('PGREMUNERATION',Tob_Temp2.GetValue('PGP_PGVALCHAMP'),False) + ' (base) : ' + FormatFloat('#,##0.00',Result);
            end;
            if (Tob_Temp.GetValue('PGP_PGVALCHAMP') = 'TAU') then
            begin
              Result := CalcZoneBasTauCoeMnt('PHB_TAUXREM',Etab,Salarie,Tob_Temp2.GetValue('PGP_PGVALCHAMP'),NumZone,DateDeb,DateFin);
              Result := RechDom('PGREMUNERATION',Tob_Temp2.GetValue('PGP_PGVALCHAMP'),False) + ' (taux) : ' + FormatFloat('#,##0.0000',Result);
            end;
            if (Tob_Temp.GetValue('PGP_PGVALCHAMP') = 'COE') then
            begin
              Result := CalcZoneBasTauCoeMnt('PHB_COEFFREM',Etab,Salarie,Tob_Temp2.GetValue('PGP_PGVALCHAMP'),NumZone,DateDeb,DateFin);
              Result := RechDom('PGREMUNERATION',Tob_Temp2.GetValue('PGP_PGVALCHAMP'),False) + ' (coefficient) : ' + FormatFloat('#,##0.00',Result);
            end;
            if (Tob_Temp.GetValue('PGP_PGVALCHAMP') = 'MNT') then
            begin
              Result := CalcZoneBasTauCoeMnt('PHB_MTREM',Etab,Salarie,Tob_Temp2.GetValue('PGP_PGVALCHAMP'),NumZone,DateDeb,DateFin);
              Result := RechDom('PGREMUNERATION',Tob_Temp2.GetValue('PGP_PGVALCHAMP'),False) + ' (montant) : ' + FormatFloat('#,##0.00',Result);
            end;
          end;
        end;

        // -------- RESULTAT --------
        if (Tob_Temp.GetValue('PGP_PGVALCHAMP') = 'RES') then
        begin
          // Récupération de la valeur de la zone à gauche de l'opérateur
          NomChamp := 'CBVALZONE' + NumZone;
          Tob_Temp2 := TPGContexte.GetCurrent.GblTob_ParamPop.FindFirst(['PGP_PGNOMCHAMP'], [NomChamp], False);
          if assigned(Tob_Temp2) then
            Resultat1 := TPGContexte.GetCurrent.Tab_ZonesCalculees[StrToInt(Tob_Temp2.GetValue('PGP_PGVALCHAMP'))];

          // Récupération de la valeur de la zone à droite de l'opérateur
          NomChamp := 'CBVALZONE' + NumZone + 'B';
          Tob_Temp2 := TPGContexte.GetCurrent.GblTob_ParamPop.FindFirst(['PGP_PGNOMCHAMP'], [NomChamp], False);
          if assigned(Tob_Temp2) then
            Resultat2 := TPGContexte.GetCurrent.Tab_ZonesCalculees[StrToInt(Tob_Temp2.GetValue('PGP_PGVALCHAMP'))];

          // Opération
          NomChamp := 'OPERATEUR' + NumZone;
          Tob_Temp2 := TPGContexte.GetCurrent.GblTob_ParamPop.FindFirst(['PGP_PGNOMCHAMP'], [NomChamp], False);
          if assigned(Tob_Temp2) then
          begin
            if (Tob_Temp2.GetValue('PGP_PGVALCHAMP') = '+') then
              Resultat := Resultat1 + Resultat2;
            if (Tob_Temp2.GetValue('PGP_PGVALCHAMP') = '-') then
              Resultat := Resultat1 - Resultat2;
            if (Tob_Temp2.GetValue('PGP_PGVALCHAMP') = '*') then
              Resultat := Resultat1 * Resultat2;
            if (Tob_Temp2.GetValue('PGP_PGVALCHAMP') = '/') then
              if (Resultat2 <> 0) then
                Resultat := Resultat1 / Resultat2;
          end;
          Result := FormatFloat('#,##0.00',Resultat);
        end;
      end;
    end;
  end;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : FC
Créé le ...... : 22/02/2007
Modifié le ... :   /  /
Description .. : Function @ de l'édition des bulletins
Mots clefs ... : PAIE;EDITION
*****************************************************************}
procedure FnPaieCreateBull(DateDeb, DateFin: TDateTime; Salarie, Etab, Page, Population: string;DerPage:String); //PT5
var
  W_PgParametre,W_TypeParametre : String;
  QTemp,QParamAssoc,Q       :TQuery;
  i,j           :integer;
  DTcloturePi, DTDebutPi, Tdate: TDateTime;
  EnAnnee, EnMois, EnJour, RMois: Word;
  CodeAssocie : String;
  CpAn, CpMs, CpJr : Word;

  QOrg, QNiv, QMontEx, QRechCumul: TQuery;
  CpAnnee, CpMois, CpJour: Word;
  CumulDD: TDateTime;
  Raz, PclCumul, st, WhereCumMois, StCum, EtabOrg: string;
  TOrg: TOB;
  CodeCalcul: string;
  Taux, Coeff, Multiplicateur: double;
begin
  TPGContexte.GetCurrent.organisme := '';
  TPGContexte.GetCurrent.niveau := '';
  TPGContexte.GetCurrent.cumul21 := '';

  if (Datefin < 1) or (DateDeb < 1) or (Salarie = '') or (Etab = '') then exit;

  if (Page = '1') or (Page = '') or (DateDeb <> TPGContexte.GetCurrent.MemDateDeb) or (DateFin <> TPGContexte.GetCurrent.MemDateFin) or (Etab <> TPGContexte.GetCurrent.MemEtab) or (DerPage='X') then
  begin
    TPGContexte.GetCurrent.MemDateDeb := DateDeb; // rupture sur les dates donc rechargement
    TPGContexte.GetCurrent.MemDateFin := DateFin;
    TPGContexte.GetCurrent.MemEtab := Etab;

    //init l'organisme associé à l'établissement du salarie
    if Assigned(TPGContexte.GetCurrent.Tob_OrganismeBull) then
    begin
      TPGContexte.GetCurrent.Tob_OrganismeBull.Free;
      TPGContexte.GetCurrent.Tob_OrganismeBull := nil;
    end;
{    QOrg := OpenSql('SELECT PSA_ETABLISSEMENT,PSA_SALARIE,PSA_TYPEDITORG,PSA_EDITORG,ETB_EDITORG ' +
      'FROM SALARIES LEFT JOIN ETABCOMPL ON PSA_ETABLISSEMENT=ETB_ETABLISSEMENT ' +
      'WHERE PSA_SALARIE IN (SELECT PPU_SALARIE FROM PAIEENCOURS ' +
      'WHERE PPU_DATEFIN>="' + UsDateTime(DebutDeMois(DateDeb)) + '" ' +
      'AND PPU_DATEFIN<="' + UsDateTime(FinDeMois(DateFin)) + '")', True);
    if not QOrg.eof then
    begin
      TPGContexte.GetCurrent.Tob_OrganismeBull := Tob.create('ORGANISMEPAIE', nil, -1);
      TPGContexte.GetCurrent.Tob_OrganismeBull.LoadDetailDB('ORGANISMEPAIE', '', '', QOrg, FALSE);
    end
    else TPGContexte.GetCurrent.Tob_OrganismeBull := nil;
    Ferme(QOrg);}
    st := 'SELECT PSA_ETABLISSEMENT,PSA_SALARIE,PSA_TYPEDITORG,PSA_EDITORG,ETB_EDITORG ' +
      'FROM SALARIES LEFT JOIN ETABCOMPL ON PSA_ETABLISSEMENT=ETB_ETABLISSEMENT ' +
      'WHERE PSA_SALARIE IN (SELECT PPU_SALARIE FROM PAIEENCOURS ' +
      'WHERE PPU_DATEFIN>="' + UsDateTime(DebutDeMois(DateDeb)) + '" ' +
      'AND PPU_DATEFIN<="' + UsDateTime(FinDeMois(DateFin)) + '")';
    TPGContexte.GetCurrent.Tob_OrganismeBull := Tob.Create('ORGANISMEPAIE',nil,-1); // PT109
    TPGContexte.GetCurrent.Tob_OrganismeBull.LoadDetailDbFromSQL('ORGANISMEPAIE',st) ;

    //init date debut exercice social pour l'alimentation des cumuls
{    QTemp := OpenSql('SELECT PEX_DATEDEBUT,PEX_DATEFIN FROM EXERSOCIAL ORDER BY PEX_DATEFIN',True);
    if not QTemp.eof then //PORTAGECWAS
      Begin
      TPGContexte.GetCurrent.Tob_ExerciceSocial := Tob.Create('Les exercices social',nil,-1);
      TPGContexte.GetCurrent.Tob_ExerciceSocial.LoadDetailDB('','','',QTemp,False);
      End
    else
      TPGContexte.GetCurrent.Tob_ExerciceSocial := nil;
    Ferme(QTemp);}
    st := 'SELECT PEX_DATEDEBUT,PEX_DATEFIN FROM EXERSOCIAL ORDER BY PEX_DATEFIN';
    TPGContexte.GetCurrent.Tob_ExerciceSocial := Tob.Create('Les exercices social',nil,-1); // PT109
    TPGContexte.GetCurrent.Tob_ExerciceSocial.LoadDetailDbFromSQL('EXERSOCIAL',st) ;

    //init FDateCp pour l'alimentation des cumuls
    TPGContexte.GetCurrent.Tob_DateClotureCp := Tob.create('Date de cloture', nil, -1);
    if TPGContexte.GetCurrent.Tob_DateClotureCp <> nil then
    begin
      QTemp := OpenSql('SELECT ETB_ETABLISSEMENT,ETB_DATECLOTURECPN FROM ETABCOMPL ', True);
      while not QTemp.eof do
      begin
        TPGContexte.GetCurrent.Tob_DateCp := Tob.create('Date de cloture', TPGContexte.GetCurrent.Tob_DateClotureCp, -1);
        if TPGContexte.GetCurrent.Tob_DateCp = nil then break;
        TPGContexte.GetCurrent.Tob_DateCp.AddChampSup('ETABLISSEMENT', False);
        TPGContexte.GetCurrent.Tob_DateCp.AddChampSup('DATECLOTURE', False);
        TDate := QTemp.FindField('ETB_DATECLOTURECPN').AsDateTime;
        if (TDate > 0) then
        begin
          RMois := 1;
          RendPeriode(DTcloturePi, DTDebutPi, TDate, DateFin);
          DecodeDate(DateFin, EnAnnee, EnMois, EnJour);
          DecodeDate(DTcloturePi, CpAn, CpMs, CpJr);
          TPGContexte.GetCurrent.CpAnnee := CpAn;
          TPGContexte.GetCurrent.CpMois := CpMs;
          TPGContexte.GetCurrent.CpJour := CpJR;
          if EnMois <= TPGContexte.GetCurrent.CpMois then
          begin
            EnAnnee := EnAnnee - 1;
            RMois := TPGContexte.GetCurrent.CpMois + 1;
          end;
          if EnMois > TPGContexte.GetCurrent.CpMois then RMois := TPGContexte.GetCurrent.CpMois + 1;
          if TPGContexte.GetCurrent.CpMois = 12 then RMois := 1;
          TPGContexte.GetCurrent.DateCp := EncodeDate(EnAnnee, RMois, 1);
        end
        else TPGContexte.GetCurrent.DateCp := idate1900;
        TPGContexte.GetCurrent.Tob_DateCp.PutValue('ETABLISSEMENT', QTemp.FindField('ETB_ETABLISSEMENT').AsString);
        TPGContexte.GetCurrent.Tob_DateCp.PutValue('DATECLOTURE', TPGContexte.GetCurrent.DateCp);
        QTemp.next;
      end;
      Ferme(QTemp);
    end;

    //Cumuls choisis ds paamsoc et à editer
    TPGContexte.GetCurrent.WhereCumul := '';
    if GetParamSocSecur('SO_PGCUMUL01','') <> '' then TPGContexte.GetCurrent.WhereCumul := ' OR PHC_CUMULPAIE="' + GetParamSocSecur('SO_PGCUMUL01','') + '" ';
    if GetParamSocSecur('SO_PGCUMUL02','') <> '' then TPGContexte.GetCurrent.WhereCumul := TPGContexte.GetCurrent.WhereCumul + ' OR PHC_CUMULPAIE="' + GetParamSocSecur('SO_PGCUMUL02','') + '" ';
    if GetParamSocSecur('SO_PGCUMUL03','') <> '' then TPGContexte.GetCurrent.WhereCumul := TPGContexte.GetCurrent.WhereCumul + ' OR PHC_CUMULPAIE="' + GetParamSocSecur('SO_PGCUMUL03','') + '" ';
    if GetParamSocSecur('SO_PGCUMUL04','') <> '' then TPGContexte.GetCurrent.WhereCumul := TPGContexte.GetCurrent.WhereCumul + ' OR PHC_CUMULPAIE="' + GetParamSocSecur('SO_PGCUMUL04','') + '" ';
    if GetParamSocSecur('SO_PGCUMUL05','') <> '' then TPGContexte.GetCurrent.WhereCumul := TPGContexte.GetCurrent.WhereCumul + ' OR PHC_CUMULPAIE="' + GetParamSocSecur('SO_PGCUMUL05','') + '" ';
    if GetParamSocSecur('SO_PGCUMUL06','') <> '' then TPGContexte.GetCurrent.WhereCumul := TPGContexte.GetCurrent.WhereCumul + ' OR PHC_CUMULPAIE="' + GetParamSocSecur('SO_PGCUMUL06','') + '" ';
    if GetParamSocSecur('SO_PGCGACQ1','') <> '' then TPGContexte.GetCurrent.WhereCumul := TPGContexte.GetCurrent.WhereCumul + ' OR PHC_CUMULPAIE="' + GetParamSocSecur('SO_PGCGACQ1','') + '" ';
    if GetParamSocSecur('SO_PGCGPRIS1','') <> '' then TPGContexte.GetCurrent.WhereCumul := TPGContexte.GetCurrent.WhereCumul + ' OR PHC_CUMULPAIE="' + GetParamSocSecur('SO_PGCGPRIS1','') + '" ';
    if GetParamSocSecur('SO_PGCGACQ2','') <> '' then TPGContexte.GetCurrent.WhereCumul := TPGContexte.GetCurrent.WhereCumul + ' OR PHC_CUMULPAIE="' + GetParamSocSecur('SO_PGCGACQ2','') + '" ';
    if GetParamSocSecur('SO_PGCGPRIS2','') <> '' then TPGContexte.GetCurrent.WhereCumul := TPGContexte.GetCurrent.WhereCumul + ' OR PHC_CUMULPAIE="' + GetParamSocSecur('SO_PGCGPRIS2','') + '" ';
    if GetParamSocSecur('SO_PGCGACQ3','') <> '' then TPGContexte.GetCurrent.WhereCumul := TPGContexte.GetCurrent.WhereCumul + ' OR PHC_CUMULPAIE="' + GetParamSocSecur('SO_PGCGACQ3','') + '" ';
    if GetParamSocSecur('SO_PGCGPRIS3','') <> '' then TPGContexte.GetCurrent.WhereCumul := TPGContexte.GetCurrent.WhereCumul + ' OR PHC_CUMULPAIE="' + GetParamSocSecur('SO_PGCGPRIS3','') + '" ';
    if GetParamSocSecur('SO_PGCGACQ4','') <> '' then TPGContexte.GetCurrent.WhereCumul := TPGContexte.GetCurrent.WhereCumul + ' OR PHC_CUMULPAIE="' + GetParamSocSecur('SO_PGCGACQ4','') + '" ';
    if GetParamSocSecur('SO_PGCGPRIS4','') <> '' then TPGContexte.GetCurrent.WhereCumul := TPGContexte.GetCurrent.WhereCumul + ' OR PHC_CUMULPAIE="' + GetParamSocSecur('SO_PGCGPRIS4','') + '" ';
  end;

  //init l'organisme associé à l'établissement du salarie
  if TPGContexte.GetCurrent.Tob_OrganismeBull <> nil then
  begin
    TOrg := TPGContexte.GetCurrent.Tob_OrganismeBull.FindFirst(['PSA_SALARIE'], [Salarie], False);
    if TOrg <> nil then
    begin
      if (V_PGI.Driver in [dbORACLE7, dbORACLE8, dbORACLE9, dbORACLE10]) then EtabOrg := Etab
      else
      begin
        EtabOrg := StringOfchar(' ', 3);
        for i := 1 to Length(Etab) do EtabOrg[i] := Etab[i];
      end;

      if (TOrg.GetValue('PSA_TYPEDITORG') = 'ETB') and (TOrg.GetValue('ETB_EDITORG') <> null) then
        TPGContexte.GetCurrent.organisme := RechDom('PGORGAFFILIATION', EtabOrg + TOrg.GetValue('ETB_EDITORG'), False)
      else
        if (TOrg.GetValue('PSA_TYPEDITORG') = 'PER') and (TOrg.GetValue('PSA_EDITORG') <> null) then
        TPGContexte.GetCurrent.organisme := RechDom('PGORGAFFILIATION', EtabOrg + TOrg.GetValue('PSA_EDITORG'), False)
      else
        TPGContexte.GetCurrent.organisme := '';
        
    end;
    if (TPGContexte.GetCurrent.organisme = '') or (TPGContexte.GetCurrent.organisme = 'Error') then
      TPGContexte.GetCurrent.organisme := RechDom('PGORGAFFILIATION', EtabOrg + '001', False);
    if (TPGContexte.GetCurrent.organisme = '') or (TPGContexte.GetCurrent.organisme = 'Error') then
      TPGContexte.GetCurrent.organisme := RechDom('PGORGAFFILIATION', Etab + '001', False);
  end;

  //Emplacement net imposable
  QNiv := OpenSql('SELECT MAX(PHB_ORDREETAT) AS ORDRE FROM HISTOBULLETIN ' +
    'WHERE PHB_DATEDEBUT="' + USDateTime(DateDeb) + '" AND PHB_DATEFIN="' + USDateTime(Datefin) + '" ' +
    'AND PHB_NATURERUB<>"BAS" ' +
    'AND (PHB_MTSALARIAL+PHB_MTPATRONAL+PHB_MTREM<>0 OR PHB_RUBRIQUE like "%.%") ' +
    'AND PHB_IMPRIMABLE="X" AND PHB_ORDREETAT>0 AND PHB_ORDREETAT<7 ' +
    'AND PHB_SALARIE="' + salarie + '" AND PHB_ETABLISSEMENT="' + Etab + '"', True);
  if not QNiv.eof then //PORTAGECWAS
    TPGContexte.GetCurrent.niveau := IntToStr(QNiv.FindField('ORDRE').AsInteger)
  else
    TPGContexte.GetCurrent.niveau := '3';
  Ferme(QNiv);

  //init heures payées + Cumul du mois en cours
  WhereCumMois := '';
  if (Page = '1') or (Page = '') then
  begin
    // Réinitialiser le tableau contenant les valeurs déjà calculées
    for j := 1 to 24 do
      TPGContexte.GetCurrent.Tab_ZonesCalculees[j] := 0;

    TPGContexte.GetCurrent.MaxCum := 0;
    Q := OpenSql('SELECT COUNT (*) NBRE FROM CUMULPAIE WHERE PCL_CUMULPAIE <> "" AND ##PCL_PREDEFINI##', True);
    if not Q.EOF then
      TPGContexte.GetCurrent.MaxCum := Q.FindField('NBRE').asInteger;
    Ferme(Q);
  end;

  for i := 1 to TPGContexte.GetCurrent.MaxCum do
  begin
    if (Page = '1') or (Page = '') then TPGContexte.GetCurrent.TabCumul[i] := '';
    TPGContexte.GetCurrent.TabCumulMois[i] := 0;
    TPGContexte.GetCurrent.TabCumulEx[i] := 0;
  end;
  if (Page = '1') or (Page = '') then
  begin
    Q := OpenSql('SELECT PCL_CUMULPAIE FROM CUMULPAIE WHERE PCL_CUMULPAIE <> "" AND ##PCL_PREDEFINI##', True);
    i := 1;
    while not Q.EOF do
    begin // Boucle de chargement des numéros de cumuls
      TPGContexte.GetCurrent.TabCumul[i] := Q.FindField('PCL_CUMULPAIE').AsString;
      i := i + 1;
      Q.Next;
    end;
    Ferme(Q);
  end;

  if (GetParamSocSecur('SO_PGTYPECUMULMOIS1','',false) = 'CMO') and (GetParamSocSecur('SO_PGCUMULMOIS1','',false) <> '') then WhereCumMois := ' OR PHC_CUMULPAIE="' + GetParamSocSecur('SO_PGCUMULMOIS1','',false) + '" ';
  if (GetParamSocSecur('SO_PGTYPECUMULMOIS2','',false) = 'CMO') and (GetParamSocSecur('SO_PGCUMULMOIS2','',false) <> '') then WhereCumMois := WhereCumMois + ' OR PHC_CUMULPAIE="' + GetParamSocSecur('SO_PGCUMULMOIS2','',false) + '" ';
  if (GetParamSocSecur('SO_PGTYPECUMULMOIS3','',false) = 'CMO') and (GetParamSocSecur('SO_PGCUMULMOIS3','',false) <> '') then WhereCumMois := WhereCumMois + ' OR PHC_CUMULPAIE="' + GetParamSocSecur('SO_PGCUMULMOIS3','',false) + '" ';
  QRechCumul := OpenSql('SELECT PHC_CUMULPAIE,SUM(PHC_MONTANT) AS MONTANT FROM HISTOCUMSAL ' +
    'WHERE PHC_SALARIE="' + Salarie + '" AND PHC_ETABLISSEMENT="' + Etab + '"' +
    'AND PHC_DATEDEBUT>="' + UsDateTime(DateDeb) + '" AND PHC_DATEFIN<="' + UsDateTime(DateFin) + '" ' +
    'AND (PHC_CUMULPAIE="21" ' + WhereCumMois + ' )' +
    'GROUP BY PHC_CUMULPAIE ', TRUE);
  while not QRechCumul.Eof do
  begin
    StCum := QRechCumul.FindField('PHC_CUMULPAIE').AsString;
    if StCum = '21' then
      TPGContexte.GetCurrent.Cumul21 := FloatToStr(QRechCumul.FindField('MONTANT').AsFloat);
    for i := 1 to TPGContexte.GetCurrent.MaxCum do
    begin
      if TPGContexte.GetCurrent.TabCumul[i] = StCum then
      begin
        TPGContexte.GetCurrent.TabCumulMois[i] := QRechCumul.FindField('MONTANT').AsFloat;
        Break;
      end;
    end;
    QRechCumul.Next;
  end;
  Ferme(QRechCumul);

{$IFNDEF CPS1}
  if Assigned(TPGContexte.GetCurrent.GblTob_ParamPop) then
  begin
    TPGContexte.GetCurrent.GblTob_ParamPop := nil;
    TPGContexte.GetCurrent.GblTob_ParamPop.free;
  end;

  CodeAssocie := '';
  if Population = '' then
  begin
    Q := OpenSQL('SELECT PNA_POPULATION FROM SALARIEPOPUL '
      + ' WHERE PNA_SALARIE = "' + Salarie + '"'
      + ' AND PNA_TYPEPOP = "PAI"', True);   // Anciennement ZLB
    if not Q.Eof then
      CodeAssocie := Q.FindField('PNA_POPULATION').AsString;
    Ferme(Q);
  end
  else
    CodeAssocie := Population;

  QParamAssoc := OpenSql('SELECT PGO_PGPARAMETRE, PGO_TYPEPARAMETRE ' +
    ' FROM PGPARAMETRESASSOC ' +
    ' WHERE PGO_CODEASSOCIE = "' + CodeAssocie + '"' +
    ' AND PGO_DATEVALIDITE <= "' + UsDateTime(DateFin) + '"' +
    ' AND PGO_PGPARAMETRE LIKE "PAI%"' +  // Anciennement ZLB
    ' ORDER BY PGO_DATEVALIDITE DESC', True);
  // Chargement des paramètres propre à la population
  if not QParamAssoc.Eof then
  begin
    W_PgParametre   := QParamAssoc.FindField('PGO_PGPARAMETRE').AsString;
    W_TypeParametre := QParamAssoc.FindField('PGO_TYPEPARAMETRE').AsString;
    QTemp := OpenSql('SELECT PGP_PGNOMCHAMP, PGP_PGVALCHAMP FROM PGPARAMETRES ' +
      'WHERE ##PGP_PREDEFINI## PGP_PGPARAMETRE = "'+ W_PgParametre + '"' +
      ' AND PGP_TYPEPARAMETRE = "' + W_TypeParametre + '"', True);
    if not QTemp.eof then
    begin
      TPGContexte.GetCurrent.GblTob_ParamPop := Tob.create('PARAMPOP', nil, -1);
      TPGContexte.GetCurrent.GblTob_ParamPop.LoadDetailDB('PARAMPOP', '', '', QTemp, FALSE);
    end
    else TPGContexte.GetCurrent.GblTob_ParamPop := nil;
    Ferme(QTemp);
  end;
  Ferme(QParamAssoc);
{$ENDIF}

  CumulDD := DateDeb;

  QTemp := OpenSQl('SELECT MIN(PHC_DATEDEBUT) AS DATEDEB FROM HISTOCUMSAL ' +
    'WHERE PHC_SALARIE="' + Salarie + '"', True);
  if not QTemp.EOF then //PORTAGECWAS
    TPGContexte.GetCurrent.DateJamais := QTemp.FindField('DATEDEB').AsDateTime
  else
    TPGContexte.GetCurrent.DateJamais := idate1900;
  Ferme(QTemp);

  if TPGContexte.GetCurrent.Tob_DateClotureCp <> nil then
  begin
    TPGContexte.GetCurrent.Tob_DateCP := TPGContexte.GetCurrent.Tob_DateClotureCp.FindFirst(['ETABLISSEMENT'], [etab], False);
    if TPGContexte.GetCurrent.Tob_DateCP <> nil then
      TPGContexte.GetCurrent.DateCp := TPGContexte.GetCurrent.Tob_DateCP.GetValue('DATECLOTURE')
    else
      TPGContexte.GetCurrent.DateCp := Idate1900;
  end
  else
    TPGContexte.GetCurrent.DateCp := Idate1900;

  TPGContexte.GetCurrent.DateDebutExerSoc := idate1900;
  if TPGContexte.GetCurrent.Tob_ExerciceSocial <> nil then
  begin
    For i := 0 to TPGContexte.GetCurrent.Tob_ExerciceSocial.detail.count-1 do
    Begin
      if (TPGContexte.GetCurrent.Tob_ExerciceSocial.detail[i].GetValue('PEX_DATEDEBUT') <= DateDeb)
      and (TPGContexte.GetCurrent.Tob_ExerciceSocial.detail[i].GetValue('PEX_DATEFIN') >= DateFin)  then
      Begin
        TPGContexte.GetCurrent.DateDebutExerSoc :=  TPGContexte.GetCurrent.Tob_ExerciceSocial.detail[i].GetValue('PEX_DATEDEBUT');
        Break;
        End;
    End;
  end;

  Raz := '';
  //Alimente les cumuls en fonction de sa periode de remise à zero
  QRechCumul := OpenSql('SELECT DISTINCT PHC_CUMULPAIE,PCL_RAZCUMUL FROM HISTOCUMSAL  ' +
    'LEFT JOIN CUMULPAIE ON PCL_CUMULPAIE=PHC_CUMULPAIE ' +
    'WHERE ##PCL_PREDEFINI## PHC_SALARIE="' + salarie + '" ' +
    'AND (PHC_CUMULPAIE="01" OR PHC_CUMULPAIE="07" OR PHC_CUMULPAIE="09" ' +
    'OR PHC_CUMULPAIE="20" OR PHC_CUMULPAIE="30" OR PHC_CUMULPAIE="40" ' + TPGContexte.GetCurrent.WhereCumul + ' ) ' +
    'ORDER BY PHC_CUMULPAIE', True);
  while not QRechCumul.Eof do
  begin
    PclCumul := QRechCumul.FindField('PHC_CUMULPAIE').AsString;
    Raz := QRechCumul.FindField('PCL_RAZCUMUL').AsString;
    if Raz = '00' then
        begin
        CumulDD := TPGContexte.GetCurrent.DateDebutExerSoc;
        end;
     if (raz = '01') or (raz = '02') or (raz = '03') or (raz = '04') or (raz = '05') or (raz = '06') or (raz = '07') or (raz = '08') or (raz = '09') or (raz = '10') or (raz = '11') or (raz = '12') then
        begin
          DecodeDate(Datedeb, CpAnnee, CpMois, CpJour);
          if CpMois < StrtoInt(raz) then Cpannee := Cpannee - 1;
          CumulDD := EncodeDate(CpAnnee, StrToInt(raz), 1);
        end;
     if raz = '99' then cumulDD := TPGContexte.GetCurrent.DateJamais;
     if raz = '20' then cumulDD := TPGContexte.GetCurrent.DateCP;
     if (Datefin > 0) and (CumulDD > 0) and (Salarie <> '') and (PclCumul <> '') then
        begin
          st := 'SELECT SUM(PHC_MONTANT) AS MONTANT,PHC_CUMULPAIE ' +
            'FROM HISTOCUMSAL WHERE PHC_SALARIE="' + Salarie + '" AND ' +
            'PHC_CUMULPAIE="' + PclCumul + '" ' +
            'AND PHC_DATEDEBUT>="' + UsDateTime(CumulDD) + '" ' +
            'AND PHC_DATEFIN<="' + UsDateTime(DateFin) + '" ' +
            'GROUP BY PHC_CUMULPAIE';
          QMontEx := OpenSql(st, TRUE);
          if not QMontEx.eof then //PORTAGECWAS
          begin
            for i := 1 to TPGContexte.GetCurrent.MaxCum do
            begin
              if TPGContexte.GetCurrent.TabCumul[i] = PclCumul then
              begin
                TPGContexte.GetCurrent.TabCumulEx[i] := QMontEx.FindField('MONTANT').AsFloat;
                Break;
              end;
            end;
          end;
          Ferme(QMontEx);
        end;
     QRechCumul.Next;
  end;
  ferme(QRechCumul);

  //Init de la zone taux rem    affectation du multiplicateur
  st := 'SELECT PHB_NATURERUB,PHB_RUBRIQUE,PHB_COTREGUL,PHB_TAUXREM,PHB_COEFFREM,PRM_CODECALCUL ' +  //PT2
    'FROM HISTOBULLETIN ' +
    'LEFT JOIN REMUNERATION ON PHB_NATURERUB=PRM_NATURERUB AND ##PRM_PREDEFINI## PHB_RUBRIQUE=PRM_RUBRIQUE ' +
    'WHERE PHB_SALARIE="' + Salarie + '" AND PHB_ETABLISSEMENT="' + Etab + '"' +
    'AND PHB_DATEDEBUT>="' + UsDateTime(DateDeb) + '" AND PHB_DATEFIN<="' + UsDateTime(DateFin) + '" ' +
    'AND PHB_NATURERUB="AAA" AND PHB_IMPRIMABLE="X" AND PRM_CODECALCUL<>"01" ' +
    'AND (PHB_TAUXREMIMPRIM="X" OR PHB_COEFFREMIMPRIM="X") ' +
    'ORDER BY PHB_ORDREETAT,PHB_RUBRIQUE ';

  TPGContexte.GetCurrent.Tob_TauxBull := Tob.create('Le taux calculé', nil, -1);
  TPGContexte.GetCurrent.Tob_TauxBull.LoadDetailDBFROMSQL ('HISTO_BULLETIN',St);

  for i := 0 to TPGContexte.GetCurrent.Tob_TauxBull.detail.count - 1 do
  begin
    TPGContexte.GetCurrent.Tob_TauxBull.detail[i].AddChampSup('MULTIPLICATEUR', False);
    CodeCalcul := TPGContexte.GetCurrent.Tob_TauxBull.detail[i].GetValue('PRM_CODECALCUL');
    Taux := TPGContexte.GetCurrent.Tob_TauxBull.detail[i].GetValue('PHB_TAUXREM');
    Coeff := TPGContexte.GetCurrent.Tob_TauxBull.detail[i].GetValue('PHB_COEFFREM');
    Multiplicateur := 0;
    if CodeCalcul = '02' then Multiplicateur := Taux * Coeff
    else if CodeCalcul = '03' then Multiplicateur := Taux / 100 * Coeff
    else if CodeCalcul = '04' then Multiplicateur := Taux
    else if CodeCalcul = '05' then Multiplicateur := Taux
    else if CodeCalcul = '08' then Multiplicateur := Coeff;
    if Coeff <> 0 then
    begin
      if CodeCalcul = '06' then Multiplicateur := Taux / Coeff
      else if CodeCalcul = '07' then Multiplicateur := Taux / 100 / Coeff;
    end;
    TPGContexte.GetCurrent.Tob_TauxBull.detail[i].PutValue('MULTIPLICATEUR', Multiplicateur);
  end;
end;

// Fonction de récupération d'une base, d'un taux, d'un coefficient ou d'un montant de rémunération
function CalcZoneBasTauCoeMnt(ChampPHB,Etab,Salarie,ValChamp,NumZone:String;DateDeb,DateFin:TDateTime):double;
var
  Requete : String;
  Q:TQuery;
begin
  Result := 0;
  Requete := 'SELECT ' + ChampPHB +
    ' FROM HISTOBULLETIN ' +
    ' WHERE PHB_NATURERUB="AAA" ' +
    ' AND PHB_ETABLISSEMENT = "' + Etab + '"' +
    ' AND PHB_SALARIE = "' + Salarie + '" ' +
    ' AND PHB_DATEDEBUT >= "' + UsDateTime(DateDeb) + '" ' +
    ' AND PHB_DATEFIN <= "' + UsDateTime(DateFin) + '" ' +
    ' AND PHB_RUBRIQUE = "' + ValChamp + '"';

  Q := OpenSql(Requete, TRUE);
  if not Q.eof then //PORTAGECWAS
  begin
    Result := Q.FindField(ChampPHB).AsFloat;
  end;
  Ferme(Q);

  TPGContexte.GetCurrent.Tab_ZonesCalculees[strtoint(NumZone)] := Result;
end;

function PGColleZeroDevant(Nombre, Long: integer): string;
begin
  result := Format('%-' + IntToStr(Long) + '.' + IntToStr(Long) + 'd', [Nombre])
end;

{function ReadParam(var St: string): string;
var
  i: Integer;
begin
  i := Pos(';', St);
  if i <= 0 then
    i := Length(St) + 1;
  result := Copy(St, 1, i - 1);
  Delete(St, 1, i);
end;}

end.


