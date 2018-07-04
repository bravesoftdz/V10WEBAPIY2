{***********UNITE*************************************************
Auteur  ...... :
Créé le ...... : 13/06/2006
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : UTOFBTAFFAPPEL ()
Mots clefs ... : TOF;UTOFBTAFFAPPEL
*****************************************************************}
Unit UTofBtAffAppel ;

Interface

Uses StdCtrls,
     Controls,
     Classes,
    {$IFDEF EAGLCLIENT}
      MaineAGL, UtilEagl,
      eMul,
    {$ELSE}
      db,
      {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}
      Fe_main,
      EdtREtat,
      mul,
    {$ENDIF}
     forms,
     sysutils,
     ComCtrls,
     HCtrls,
     HTB97,
     HEnt1,
     HMsgBox,
     HeureUtil,
     UTOF ,
     UTob,
     AglInit,
     UtilsMail
     ;

Type
  TOF_BTAFFAPPEL = Class (TOF)
    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    //procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnDisplay                ; override ;
    procedure OnClose                  ; override ;
    procedure OnCancel                 ; override ;

Private

  NOMTIERS   : THEdit;
  ADRESSE1   : THEdit;
  ADRESSE2   : THEdit;
  ADRESSE3   : THEdit;
  CODEPOSTAL : THEdit;
  VILLE      : THEdit;
  NOMCONTACT : THEdit;
  TELEPHONE  : THEdit;
  CONTRAT    : THEdit;
  AFFAIRE    : THEdit;
  DATEFIN    : THEdit;
  DATEDEB    : THEdit;
  HEUREDEB   : THEdit;
  HEUREFIN   : THEdit;
  TNbHeure   : THLabel;
  NbHeure    : THEdit;

  BTHAUTE    : TToolbarButton97;
  BTBASSE    : TToolbarButton97;
  BTNORMALE  : TToolbarButton97;

  BTValider  : TToolbarButton97;
  BTMail     : TToolbarButton97;

  LAFFAIRE   : THLabel;

  DESCRIPTIF : TMemo;

  GRBTYPEACTION : TGroupBox;

  GRIDAFFECTATION : THGrid;

  CodeAppel     : String;
  CodeRessource	: String;
  Equipe        : String;
  RefEquipe     : String;
  NumEvent      : Integer;

  OldDDeb       : TDateTime;
  OldDFin       : TDateTime;
  OldHDeb       : TTime;
  OldHFin       : TTime;

  procedure ChargeAppel;
  
  procedure DateDebOnExit(Sender: TObject);
  procedure DateFinOnExit(Sender: TObject);

  procedure GridAffectation_OnClick(Sender: TObject);

  procedure HeureDebOnExit(Sender: TObject);
  procedure HeureFinOnExit(Sender: TObject);

  procedure BTMailOnClick(Sender: Tobject);
  procedure BTValiderOnClick(Sender: Tobject);

  end ;

Implementation

uses uBtpChargeTob,UdateUtils ;

procedure TOF_BTAFFAPPEL.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_BTAFFAPPEL.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_BTAFFAPPEL.ChargeAppel;
Var TobTempo : Tob;
    Contact  : Integer;
    I        : Integer;
    Tiers    : String;
    StArg    : String;
    QQ			 : TQuery;
begin

  Contact := 0;

  //chargement de la grille des affectations
  GridAffectation.Cells[0, 0]   := 'Affectation';
  GridAffectation.ColLengths[0] := GridAffectation.Width;

  BTHaute.Visible   := False;
  BTNormale.Visible := False;
  BTBasse.Visible   := False;

  // Chargement de la tob Adresse
  TobTempo := TOB.Create ('Les Adresses', Nil, -1);
  StArg := 'SELECT ADR_LIBELLE, ADR_LIBELLE2, ADR_ADRESSE1,ADR_ADRESSE2,ADR_ADRESSE3, ADR_CODEPOSTAL, ADR_VILLE, ADR_NUMEROCONTACT';
  StArg := StArg + ' FROM ADRESSES WHERE ADR_REFCODE="' + CodeAppel + '"';
  QQ := OpenSql(StArg, True);
  TobTempo.LoadDetailDB('ADRESSES', '', '', QQ, False);

  If TobTempo.detail.count > 0 then
  Begin
    Contact := TobTempo.Detail[0].GetValue('ADR_NUMEROCONTACT');
    NomTiers.text   := TobTempo.Detail[0].GetValue('ADR_LIBELLE') + ' ' + TobTempo.Detail[0].GetValue('ADR_LIBELLE2');
    Adresse1.text   := TobTempo.Detail[0].GetValue('ADR_ADRESSE1');
    Adresse2.text   := TobTempo.Detail[0].GetValue('ADR_ADRESSE2');
    Adresse3.text   := TobTempo.Detail[0].GetValue('ADR_ADRESSE3');
    CodePostal.text := TobTempo.Detail[0].GetValue('ADR_CODEPOSTAL');
    Ville.text      := TobTempo.Detail[0].GetValue('ADR_VILLE');
  end;

  TobTempo.free;

  // Chargement de la tob Affaire
  StArg := 'SELECT AFF_AFFAIRE1, AFF_AFFAIRE2, AFF_AFFAIRE3, AFF_TIERS, ';
  StArg := StArg + 'AFF_DESCRIPTIF, AFF_CHANTIER, AFF_AFFAIREINIT, AFF_PRIOCONTRAT ';
  StArg := StArg + 'FROM AFFAIRE WHERE AFF_AFFAIRE="' + CodeAppel + '"';
  TobTempo := TOB.Create ('Les Affaires', Nil, -1);
  TobTempo.LoadDetailFromSQL(StArg);
  If TobTempo.detail.count > 0 then
  Begin
     LAffaire.caption := 'Appel N°' + TobTempo.Detail[0].GetValue('AFF_AFFAIRE1') + TobTempo.Detail[0].GetValue('AFF_AFFAIRE2') + TobTempo.Detail[0].GetValue('AFF_AFFAIRE3');
     Tiers            := TobTempo.Detail[0].GetValue('AFF_TIERS');
     Contrat.Text     := TobTempo.Detail[0].GetValue('AFF_AFFAIREINIT');
     Affaire.Text     := TobTempo.Detail[0].GetValue('AFF_CHANTIER');
     Descriptif.Text  := TobTempo.Detail[0].GetValue('AFF_DESCRIPTIF');
     if TobTempo.Detail[0].GetValue('AFF_PRIOCONTRAT') = 1 then
        BTHaute.Visible := True
     Else if TobTempo.Detail[0].GetValue('AFF_PRIOCONTRAT') = 2 then
        BTNormale.Visible := True
     Else if TobTempo.Detail[0].GetValue('AFF_PRIOCONTRAT') = 3 then
       BTBasse.Visible := True;
  end;
  TobTempo.free;

  // Chargement de la tob Contact
  if Contact <> 0 then
  begin
    StArg := 'SELECT C_NOM, C_PRENOM, C_TELEPHONE FROM CONTACT WHERE C_TIERS="' + Tiers + '" AND C_NUMEROCONTACT=' + IntToStr(Contact);
    TobTempo := TOB.Create ('Les Contacts', Nil, -1);
    TobTempo.LoadDetailFromSQL(StArg);
    If TobTempo.detail.count > 0 then
    Begin
      NomContact.Text := TobTempo.Detail[0].GetValue('C_PRENOM') + ' ' + TobTempo.Detail[0].GetString('C_NOM');
      Telephone.Text  := TobTempo.Detail[0].GetValue('C_TELEPHONE');
    end;
    TobTempo.free;
  end;

  //Chargement de la grille affectation
  StArg := 'SELECT ARS_RESSOURCE, ARS_LIBELLE FROM BTEVENPLAN LEFT JOIN RESSOURCE ON BEP_RESSOURCE = ARS_RESSOURCE WHERE BEP_AFFAIRE = "' + CodeAppel + '"';
  TobTempo := TOB.Create ('Affectation', Nil, -1);
  TobTempo.LoadDetailFromSQL (StArg);
  If Assigned(TobTempo) then
  Begin
    TobTempo.PutGridDetail(GRIDAFFECTATION, False, False, 'ARS_LIBELLE');
  end;

  for i := 0 to TobTempo.Detail.count - 1 do
  begin
    if TobTempo.detail[I].GetValue('ARS_RESSOURCE') = CodeRessource then
    Begin
      GridAffectation.Row := I+1;
      Break;
    end;
  end;

  TobTempo.free;

  GridAffectation_Onclick(GridAffectation);

  DateDeb.Enabled   := True;
  DateFin.Enabled   := True;
  HeureDeb.Enabled  := True;
  HeureFin.Enabled  := True;

end ;

procedure TOF_BTAFFAPPEL.OnArgument (S : String ) ;
begin
  Inherited ;

  CodeAppel     := Trim(ReadTokenSt(S));;
  CodeRessource := Trim(ReadTokenSt(S));

  //Déclaration des zones écran
  NomTiers   := THEdit(GetControl('NOMTIERS'));
  Adresse1   := THEdit(GetControl('ADRESSE1'));
  Adresse2   := THEdit(GetControl('ADRESSE2'));
  Adresse3   := THEdit(GetControl('ADRESSE3'));
  CodePostal := THEdit(GetControl('CODEPOSTAL'));
  Ville      := THEdit(GetControl('VILLE'));

  NomContact := THEdit(GetControl('NOMCONTACT'));
  Telephone  := THEdit(GetControl('TELEPHONE'));
  Contrat    := THEdit(GetControl('CONTRAT'));
  Affaire    := THEdit(GetControl('AFFAIRE'));

  DateDeb    := THEdit(GetControl('DATEDEB'));
  DateFin    := THEdit(GetControl('DATEFIN'));
  HeureDeb   := THEdit(GetControl('HEUREDEB'));
  HeureFin   := THEdit(GetControl('HEUREFIN'));
  TNbHeure   := THLabel(Getcontrol('TNbHeure'));
  NbHeure    := THEdit(Getcontrol('DUREE'));

  LAffaire   := THLabel(GetControl('LAFFAIRE'));

  BTHaute    := TToolbarButton97(GetControl('BTHAUTE'));
  BTNormale  := TToolbarButton97(GetControl('BTNORMALE'));
  BTBasse    := TToolbarButton97(GetControl('BTBASSE'));

  BTValider  := TToolBarButton97(GetControl('BValider'));
  BTValider.OnClick := BtValiderOnClick;

  BTMail     := TToolBarButton97(GetControl('BTMAIL'));
  BTMail.OnClick := BtMailOnClick;

  GRBTYPEACTION := TGroupBox(GetControl('GRBTYPEACTION'));

  DESCRIPTIF := TMemo(GetControl('DESCRIPTIF'));

  GRIDAFFECTATION := THGrid(GetControl('GRIDAFFECTATION'));
  GridAffectation.OnClick := GridAffectation_OnClick;

  DateDeb.OnExit          := DateDebOnExit;
  DateFin.OnExit          := DateFinOnExit;
  HeureDeb.OnExit         := HeureDebOnExit;
  HeureFin.OnExit         := HeureFinOnExit;

  ChargeAppel;

end ;


//Mise à jour de l'enregistrement évènement et appel avec les nouvelles valeur date et heure
procedure TOF_BTAFFAPPEL.BTValiderOnClick(Sender : Tobject);
Var StSQL    : string;
    SqlB     : string;
    QQ       : TQuery;
    TobParam : TOB;
    TDateDeb : Tdatetime;
    TDateFin : TDateTime;
    Hdeb     : TTime;
    HFin     : TTime;
    NbMinute : Double;
begin

  //vérification si un ou plusieur évènement pour cet appel
  if CodeAppel ='' then Exit;
  
  Hdeb        := StrToTime(HEUREDEB.Text);
  Hfin        := StrToTime(HEUREFin.Text);
  TDateDeb    := StrToDate(DateDeb.Text)+HDeb;
  TDateFin    := StrToDate(DateFin.Text)+Hfin;

  Try
    TobParam := Tob.Create('AppelEvenement', nil, -1);
    //
    TobParam.AddChampSupValeur('RETOUR', 1);
    TobParam.AddChampSupValeur('ACTION', 'M');
    TobParam.AddChampSupValeur('EQUIPE', Equipe);
    //
    StSql := 'SELECT BEP_CODEEVENT FROM BTEVENPLAN WHERE BEP_AFFAIRE = "' + CodeAppel + '"';
    QQ := OpenSQL(StSQL, True,-1,'',false);
    if not QQ.eof then
    begin
      If QQ.RecordCount > 1 then
      Begin
        TheTob := TobParam;
        AglLanceFiche ('BTP','BTSUPPEVENT','','','');
        TheTob := nil;
        if TOBParam.getValue('RETOUR')= 0 then
        begin
          TOBParam.free;
          exit;
        end;
      end
      else
        NumEvent := QQ.FindField('BEP_CODEEVENT').Asinteger;
    end;
  finally
    Ferme(QQ);
  end;

  NbMinute := CalculDureeEvenement(TdateDeb, TDateFin);

	 //Traitement de Modification ou déplacement (1 - ensemble de l'affectation, 2 - Equipe, 3 - Item Sélectionnées)
   StSQL := 'UPDATE BTEVENPLAN SET BEP_DATEDEB="' + USDATETIME(TDateDeb) + '", ';
   StSQL := StSQL + 'BEP_DATEFIN="'   + USDATETIME(TDatefin) + '", ';
   StSQL := StSQL + 'BEP_HEUREDEB="'  + USTIME(HDeb) + '", ';
   StSQl := StSQl + 'BEP_HEUREFIN="'  + USTIME(HFin) + '", ';
   StSQl := StSQl + 'BEP_DUREE='      + FloatToStr(NbMinute);

  If TobParam.GetValue('RETOUR') = 1 then //L'ensemble des affectation à l'appel
  Begin
    SqlB := 'WHERE BEP_AFFAIRE ="' + CodeAppel + '"';
    StSql := StSQl +SqlB;
  end
  else if TobParam.GetValue('RETOUR') = 2 Then //toute l'équipe
  Begin
   		SqlB := 'WHERE BEP_AFFAIRE="' + CodeAppel + '" AND BEP_EQUIPERESS="'+ Equipe+ '"';
      StSQL := StSQl + SqlB;
  end
  Else If TobParam.GetValue('RETOUR') = 3 Then //Uniquement l'évènement en cours
  Begin
   		SqlB := ' WHERE BEP_CODEEVENT="' + IntToStr(NumEvent) + '"';
      StSQL := StSQl + SqlB;
  end;

  ExecuteSQL(StSql);

  // mise à jour de l'appel
  StSql := 'UPDATE AFFAIRE SET AFF_DATEREPONSE="'+ USDateTIME(TDateDeb)+'" WHERE AFF_AFFAIRE="'+ CodeAppel+'"';
  ExecuteSQL(StSql);

  TobParam.free;

end;

procedure TOF_BTAFFAPPEL.BTMailOnClick(Sender : Tobject);
begin

end;


procedure TOF_BTAFFAPPEL.GridAffectation_OnClick (Sender: TObject) ;
Var TobTempo  : Tob;
    StArg     : String;
    EcartDuree: double;
begin
  Inherited ;

  CodeRessource := GridAffectation.CellValues[0, GridAffectation.Row];

  //Chargement de la partie temps date et durée...
  StArg := 'SELECT BTA_LIBELLE, BEP_DATEDEB, BEP_HEUREDEB, BEP_DATEFIN, BEP_HEUREFIN, BEP_DUREE, AFF_CREERPAR, ';
  StArg := StArg + 'BEP_EQUIPERESS, BEP_CODEEVENT, BEP_REFEQUIPE ';
  StArg := StArg + 'FROM BTEVENPLAN LEFT JOIN RESSOURCE ON BEP_RESSOURCE = ARS_RESSOURCE ';
  StArg := StArg + 'LEFT JOIN BTETAT ON BTA_BTETAT = BEP_BTETAT ';
  StArg := StArg + 'LEFT JOIN AFFAIRE ON AFF_AFFAIRE = BEP_AFFAIRE ';
  StArg := StArg + 'WHERE BEP_AFFAIRE = "' + CodeAppel + '" AND ARS_LIBELLE = "' + CodeRessource + '"';

  TobTempo := TOB.Create ('DateTime', Nil, -1);
  TobTempo.LoadDetailFromSQL (StArg);

  Equipe        := TobTempo.Detail[0].GetValue('BEP_EQUIPERESS');
  NumEvent      := TobTempo.Detail[0].GetValue('BEP_CODEEVENT');
  RefEquipe     := TobTempo.Detail[0].GetValue('BEP_REFEQUIPE');

  DateDeb.text  := DateToStr(TobTempo.Detail[0].GetValue('BEP_DATEDEB'));
  DateFin.text  := DateToStr(TobTempo.Detail[0].GetValue('BEP_DATEFIN'));
  HeureDeb.text := FormatDateTime('hh:mm', TobTempo.Detail[0].GetValue('BEP_DATEDEB')); //TimeToStr(TobTempo.Detail[0].GetValue('BEP_DATEDEB'));
  HeureFin.text := FormatDateTime('hh:mm', TobTempo.Detail[0].GetValue('BEP_DATEFIN')); //TimeToStr(TobTempo.Detail[0].GetValue('BEP_DATEFIN'));

  OldDDeb       := StrToDate(DateDeb.Text) + StrToTime(HEUREDeb.text);
  OldDFin       := StrToDate(DateFin.Text) + StrToTime(HEUREFin.text);
  OldHDeb       := StrToTime(HeureDeb.Text);
  OldHfin       := StrToTime(HeureFin.Text);

  EcartDuree      := TobTempo.Detail[0].GetValue('BEP_DUREE');
  if EcartDuree > 59 then
    TNbHeure.Caption := 'Nb Heure(s):'
  else
    TNbHeure.Caption := 'Nb Minutes :';

  NbHeure.Text := LibelleDuree(EcartDuree, False);
                    
  GRBTYPEACTION.Caption := TobTempo.Detail[0].GetValue('BTA_LIBELLE');

  if TobTempo.Detail[0].GetValue('AFF_CREERPAR') = 'TAC' then
     Ecran.Caption := 'Fiche visite préventive sous contrat'
  Else
     Ecran.Caption := 'Fiche Appel';

  TobTempo.free;

end ;

procedure TOF_BTAFFAPPEL.OnClose ;
begin
  Inherited ;
end ;

procedure TOF_BTAFFAPPEL.OnDisplay () ;
begin
  Inherited ;
end ;

procedure TOF_BTAFFAPPEL.OnCancel () ;
begin
  Inherited ;
end ;

procedure TOF_BTAFFAPPEL.OnUpdate;
begin
  inherited;
end;

procedure TOF_BTAFFAPPEL.DateDebOnExit(Sender: TObject);
Var TDateDeb  : TDateTime;
    TDateFin  : TDateTime;
    NbMinute  : Double;
begin

  TDateDeb    := StrToDate(DateDeb.Text)+StrToTime(HeureDeb.text);
  TDateFin    := StrToDate(DateFin.Text)+StrToTime(HeureFin.text);

  if TDateDeb > TDatefin then
  begin
    PGIError('La date de début ne peut-être supérieure à la date de fin', 'Erreur Saisie');
    DateDeb.Text := DateToStr(OldDDeb);
    Datedeb.SetFocus;
    Exit;
  end;

  if TDatedeb = OldDdeb then
    exit
  else
    BtValider.Visible := True;

  NbMinute := CalculDureeEvenement(TDateDeb, TDatefin);
  if NbMinute > 59 then
    TNbHeure.Caption := 'Nb Heure(s):'
  else
    TNbHeure.Caption := 'Nb Minutes :';

  NbHeure.Text := LibelleDuree(NbMinute, False);

  OldDDeb := TDateDeb;

end;

procedure TOF_BTAFFAPPEL.DateFinOnExit(Sender: TObject);
Var TDateDeb  : TDateTime;
    TDateFin  : TDateTime;
    NbMinute : Double;
begin

  TDateDeb    := StrToDate(DateDeb.Text)+StrToTime(HeureDeb.text);
  TDateFin    := StrToDate(DateFin.Text)+StrToTime(HeureFin.text);

  if TDateFin < TDateDeb then
  begin
    PGIError('La date de fin ne peut pas être inférieure à la date de début', 'Erreur Saisie');
    DateFin.Text := DateToStr(OldDFin);
    DateFin.SetFocus;
    exit;
  end;

  if TDateFin = idate2099 then exit;

  if TDateFin = OldDFin   then
    exit
  else
    BtValider.Visible := True;

  NbMinute := CalculDureeEvenement(TDateDeb,TDateFin);
  if NbMinute > 59 then
    TNbHeure.Caption := 'Nb Heure(s):'
  else
    TNbHeure.Caption := 'Nb Minutes :';

  NbHeure.Text := LibelleDuree(NbMinute, False);

  OldDFin := TDateFin;

end;

procedure TOF_BTAFFAPPEL.HeureDebOnExit(Sender: TObject);
Var	Heure       : TdateTime;
    HeureDebAM  : TdateTime;
    HeureFinAM  : TDateTime;
    HeureDebPM  : Tdatetime;
    HeureFinPM  : TDateTime;
    TDateDeb    : TDateTime;
    TDateFin    : TDateTime;
    NBMinute    : Double;
Begin

  if copy(HeureDeb.text, 0, 2) > '24' then
  Begin
    HeureDeb.Text := TimetoStr(OldHDeb);
    Exit;
  end;

  if copy(HeureDeb.text, 4, 2) > '59' then
  Begin
    HeureDeb.Text := TimetoStr(OldHDeb);
    Exit;
  end;

  //Il faut contrôler que la date saisie soit cohérente avec les heure de début et de fin....
  //soit de la ressource soit du paramétrage...
  Heure           := StrTotime(HeureDeb.text);

  //Si jamais il y a un calendrier sur la ressource celui-ci ne sera pas pris en compte ici...
  HeureDebAm      := GetDebutMatinee;
  HeureFinAm      := GetFinMatinee;
  HeureDebPm      := GetDebutApresMidi;
  HeurefinPm      := GetfinApresMidi;

  if (Heure < heureDebAM) then Heure := heureDebAM;

  if (Heure > heureFinPM) then Heure := heureFinPM;

  if (Heure > heureFinAM) And (Heure < heureDebPM) then Heure := HeureFinAM;

  HeureDeb.Text := TimetoStr(Heure);

  TDateDeb    := StrToDate(DateDeb.Text)+StrToTime(HeureDeb.text);
  TDateFin    := StrToDate(DateFin.Text)+StrToTime(HeureFin.text);

  if TDateDeb = OldDDeb   then
    exit
  else
    BtValider.Visible := True;

  NbMinute := CalculDureeEvenement(TDateDeb,TDateFin);
  if NbMinute > 59 then
    TNbHeure.Caption := 'Nb Heure(s):'
  else
    TNbHeure.Caption := 'Nb Minutes :';

  NbHeure.Text := LibelleDuree(NbMinute, False);

  OldHDeb := StrToTime(HEUREDEB.text);

end;

procedure TOF_BTAFFAPPEL.HeureFinOnExit(Sender: TObject);
Var	Heure       : TdateTime;
    HeureDebAM  : TdateTime;
    HeureFinAM  : TDateTime;
    HeureDebPM  : Tdatetime;
    HeureFinPM  : TDateTime;
    TDateDeb    : TDateTime;
    TDateFin    : TDateTime;
    NBMinute    : Double;
begin

  if copy(HeureFin.text, 0, 2) > '24' then
  Begin
    Heurefin.Text := TimetoStr(OldHFin);
    Exit;
  end;

  if copy(Heurefin.text, 4, 2) > '59' then
  Begin
    Heurefin.Text := TimetoStr(OldHFin);
    Exit;
  end;

  //Il faut contrôler que la date saisie soit cohérente avec les heure de début et de fin....
  //soit de la ressource soit du paramétrage...
  Heure           := StrTotime(HeureFin.text);

  //Si jamais il y a un calendrier sur la ressource celui-ci ne sera pas pris en compte ici...
  HeureDebAm      := GetDebutMatinee;
  HeureFinAm      := GetFinMatinee;
  HeureDebPm      := GetDebutApresMidi;
  HeurefinPm      := GetfinApresMidi;

  if (Heure < HeureDebAM) then Heure := heureDebAM;

  if (Heure > HeureFinPM) then Heure := heureFinPM;

  if (Heure > HeureFinAM) And (Heure < heureDebPM) then Heure := heureDebPM;

  HeureFin.Text := TimetoStr(Heure);

  TDateDeb    := StrToDate(DateDeb.Text)+StrToTime(HeureDeb.text);
  TDateFin    := StrToDate(DateFin.Text)+StrToTime(HeureFin.text);

  if (TDateFin < TDateDeb) then
  begin
    Heurefin.Text := TimetoStr(OldHFin);
    Datefin.Text  := DateToStr(OldDFin);
    HeureDeb.Text := TimetoStr(OldHDeb);
    DateDeb.Text  := DateToStr(OldDDeb);
    TDateDeb      := StrToDate(DateDeb.Text)+StrToTime(HeureDeb.text);
    TDateFin      := StrToDate(DateFin.Text)+StrToTime(HeureFin.text);
  end;

  if TDateFin = OldDFin   then
    exit
  else
    BtValider.Visible := True;

  NbMinute := CalculDureeEvenement(TdateDeb, TDateFin);
  if NbMinute > 59 then
    TNbHeure.Caption := 'Nb Heure(s):'
  else
    TNbHeure.Caption := 'Nb Minutes :';

  NbHeure.Text := LibelleDuree(NbMinute, False);

  OldHFin := StrToTime(HEUREFin.text);

end;

Initialization
  registerclasses ( [ TOF_BTAFFAPPEL ] ) ;
end.

