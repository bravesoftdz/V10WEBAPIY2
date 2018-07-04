{***********UNITE*************************************************
Auteur  ...... : LS
Créé le ...... : 11/04/2007
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : BTACCEPTAFF_MUL ()
Mots clefs ... : TOF;BTACCEPTAFF_MUL
*****************************************************************}
Unit BTACCEPTAFF_MUL_TOF ;

Interface

Uses StdCtrls, 
     Controls, 
     Classes,
     M3FP,
{$IFDEF EAGLCLIENT}
     MaineAGL,emul,
     eMul,
     uTob,
     HQry,
{$ELSE}
      {$IFNDEF ERADIO}
      Fe_Main,
      {$ENDIF !ERADIO}
     db,
     {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF}
     mul,
{$ENDIF}
     forms,
     sysutils,
     ComCtrls,
     HCtrls,
     HEnt1,
     HMsgBox,
     HTB97,
     HDB,
     UTOF,
     UTOB,
     Ent1,
     utofAfBaseCodeAffaire,
     ///
     AglInit;


Type
  TOF_BTACCEPTAFF_MUL = Class (TOF_AFBASECODEAFFAIRE)
    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnDisplay                ; override ;
    procedure OnClose                  ; override ;
    procedure OnCancel                 ; override ;

    procedure NomsChampsAffaire(var Aff, Aff0, Aff1, Aff2, Aff3, Aff4, Aff_, Aff0_, Aff1_, Aff2_, Aff3_, Aff4_, Tiers, Tiers_: THEdit); override;
    
  Public

  Private

    BCherche    : TToolbarButton97;
    BOuvrir  		: TToolbarButton97;
    BOuvrir1 		: TToolbarButton97;
    BDefinitive : TToolbarButton97;

    TobRAD      : TOB;

    Fliste 			: THDbGrid;

    CodeAffaire	: ThEdit;

    TAffaire		: THLabel;
    TAffaire0		: THLabel;

    ChkModele		: TCheckBox;
    ChkAdmin		: TCheckBox;

    Responsable : THedit;

    procedure AcceptationAffaire(TheAffaire : string);
    Procedure BDefinitiveOnClick(Sender: TObject);
    procedure FlisteDblClick(Sender: TObject);
    procedure MultiAcceptation(Sender: TObject);
    procedure SaisieResteADepenser;
    procedure ValidationRAD(Affaire: String; DateArrete : TDateTime);
    procedure ResponsableOnElipsisClick(Sender: Tobject);
    procedure ControleChamp(Champ, Valeur: String);

  end ;

Implementation
Uses UtofRessource_Mul, ADODB;

procedure TOF_BTACCEPTAFF_MUL.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_BTACCEPTAFF_MUL.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_BTACCEPTAFF_MUL.OnUpdate ;
begin
  Inherited ;
end ;

procedure TOF_BTACCEPTAFF_MUL.OnLoad ;
begin
  Inherited ;
end ;

procedure TOF_BTACCEPTAFF_MUL.OnArgument (S : String ) ;
var Critere   : string;
    Valeur    : string;
    Champ     : string;
    x         : integer;
    CC        : THValCombobox;
begin
  fMulDeTraitement := true;
  Inherited ;
  fTableName := 'AFFAIRE';

  Repeat
    Critere:=uppercase(ReadTokenSt(stArgument)) ;
    valeur := '';
    if Critere<>'' then
    begin
      x:=pos('=',Critere);
      if x<>0 then
      begin
        Champ  :=copy(Critere,1,x-1);
        Valeur :=copy(Critere,x+1,length(Critere));
      end
      else
        Champ := Critere;
      ControleChamp(Champ, Valeur);
    end;
  until Critere='';

  // gestion Etablissement (BTP)
  CC:=THValComboBox(GetControl('BTBETABLISSEMENT')) ;
  if CC<>Nil then PositionneEtabUser(CC) ;

  // gestion Domaine (BTP)
  CC:=THValComboBox(GetControl('AFF_DOMAINE')) ;
  if CC<>Nil then PositionneDomaineUser(CC) ;

  //RECUPERATION DES ZONES ECRAN
  //
  CodeAffaire := THEDit(GetControl('AFF_AFFAIRE'));

  if GetControl('AFF_RESPONSABLE') <> nil then Responsable := THEDit(GetControl('AFF_RESPONSABLE'));
  //
  TAffaire 	:= THLabel(GetControl('TAFF_AFFAIRE'));
  TAffaire0 := THLabel(GetControl('TAFFAIRE0'));
  //
	ChkModele := TCheckBox(GetControl('AFF_MODELE'));
	ChkAdmin  := TCheckBox(GetControl('AFF_ADMINISTRATIF'));
  //
	Fliste := THDbGrid (GetControl('FLISTE'));
  Fliste.OnDblClick := FlisteDblClick;
  //
  Responsable.OnElipsisClick := ResponsableOnElipsisClick;
  //
  BCherche := TtoolBarButton97(GetControl('BCHERCHE'));
  //
  SetcontrolProperty('BDefinitive','Visible', False);
  //
  if assigned(TtoolBarButton97(GetControl('BOUVRIR1'))) then
  begin
    BOuvrir1 := TtoolBarButton97(GetControl('BOUVRIR1'));
    BOuvrir1.OnClick := MultiAcceptation;
  end;
  //
  if assigned(TtoolBarButton97(GetControl('BOUVRIR'))) then
  begin
    BOuvrir := TtoolBarButton97(GetControl('BOUVRIR'));
    BOuvrir.OnClick := MultiAcceptation;
  end;

  If Assigned(TtoolBarButton97(Getcontrol('BDEFINITIVE'))) Then
  Begin
    BDefinitive := TtoolBarButton97(GetControl('BDEFINITIVE'));
    BDefinitive.OnClick := BDefinitiveOnClick;
  end;

  if ecran.Name = 'BTSAISRESTDEP_MUL' then
  begin
    SetControlText('AFF_AFFAIRE0', 'A');
    SetControlText('AFFAIRE0', 'A');
    SetControlProperty('AFFAIRE0', 'Enabled', False);
    SetcontrolProperty('BOuvrir','Visible', True);
    SetcontrolProperty('BDefinitive','Visible', True);
    SetcontrolProperty('bSelectAll','Visible', True);
    Fliste.Multiselection := True;
  end;

//uniquement en line
{*  TTabSheet(GetControl('PCOMPLEMENT',true)).TabVisible := False;
  Ecran.Caption := 'Acceptation Chantier';
  TAffaire.Caption := 'Chantier :';
  TAffaire0.Caption := 'Type de Chantier';
  SetControlText('AFFAIRE0', 'A');
  SetControlProperty('AFFAIRE0', 'Enabled', False);
  ChkModele.Caption := 'Chantier Modele';
  ChkAdmin.Caption  := 'Chantier Administratif';
  UpdateCaption(TFMul(Ecran)) ;
*}

end ;

Procedure TOF_BTACCEPTAFF_MUL.ControleChamp(Champ : String;Valeur : String);
begin

  if (champ = 'STATUT') then
  begin
    if      Valeur = 'AFF' then
    begin
      if assigned(GetControl('AFF_AFFAIRE0'))  then SetControlText('AFF_AFFAIRE0', 'A')
      else if assigned(GetControl('AFFAIRE0')) then SetControlText('AFFAIRE0', 'A');
      SetControlText('XX_WHERE', ' AND SUBSTRING(GP_AFFAIRE,1,1) IN ("","A")');
      SetControlProperty('BSELECTAFF1', 'Hint', 'Recherche Chantier');
      if assigned(GetControl('TAFF_AFFAIRE'))    then SetControlText('TAFF_AFFAIRE', 'Code Chantier')
      Else if assigned(GetControl('TGPAFFAIRE')) then SetControlText('TGPAFFAIRE', 'Code Chantier');
    end
    Else if Valeur = 'APP' then
    Begin
      if assigned(GetControl('AFF_AFFAIRE0'))  then SetControlText('AFF_AFFAIRE0', 'W')
      else if assigned(GetControl('AFFAIRE0')) then SetControlText('AFFAIRE0', 'W');
      SetControlText('XX_WHERE', ' AND SUBSTRING(GP_AFFAIRE,1,1) IN ("","W")');
      SetControlProperty('BSELECTAFF1', 'Hint', 'Recherche Appel');
      if assigned(GetControl('TAFF_AFFAIRE'))    then SetControlText('TAFF_AFFAIRE', 'Code Appel')
      Else if assigned(GetControl('TGPAFFAIRE')) then SetControlText('TGPAFFAIRE', 'Code Appel');
    End
    Else if Valeur = 'INT' then
    Begin
      if assigned(GetControl('AFF_AFFAIRE0'))  then SetControlText('AFF_AFFAIRE0', 'I')
      else if assigned(GetControl('AFFAIRE0')) then SetControlText('AFFAIRE0', 'I');
      SetControlText('XX_WHERE', ' AND SUBSTRING(GP_AFFAIRE,1,1) IN ("","I")');
      SetControlProperty('BSELECTAFF1', 'Hint', 'Recherche Contrat');
      if assigned(GetControl('TAFF_AFFAIRE'))    then SetControlText('TAFF_AFFAIRE', 'Code Contrat')
      Else if assigned(GetControl('TGPAFFAIRE')) then SetControlText('TGPAFFAIRE', 'Code Contrat');
    End
    else if Valeur = 'PRO' then
    Begin
      if assigned(GetControl('AFF_AFFAIRE0'))  then SetControlText('AFF_AFFAIRE0', 'P')
      else if assigned(GetControl('AFFAIRE0')) then SetControlText('AFFAIRE0', 'P');
      SetControlText('XX_WHERE', ' AND SUBSTRING(GP_AFFAIRE,1,1) IN ("","P")');
      SetControlProperty('BSELECTAFF1', 'Hint', 'Recherche Appel d''offre');
      if assigned(GetControl('TAFF_AFFAIRE'))    then SetControlText('TAFF_AFFAIRE', 'Code Appel d''Offre')
      Else if assigned(GetControl('TGPAFFAIRE')) then SetControlText('TGPAFFAIRE', 'Code Appel d''offre');
    end
    Else
    begin
      if assigned(GetControl('AFF_AFFAIRE0'))  then SetControlText('AFF_AFFAIRE0', '')
      else if assigned(GetControl('AFFAIRE0')) then SetControlText('AFFAIRE0', '');
      SetControlProperty('BSELECTAFF1', 'Hint', 'Recherche Affaire');
      if assigned(GetControl('TAFF_AFFAIRE'))    then SetControlText('TAFF_AFFAIRE', 'Code Affaire')
      Else if assigned(GetControl('TGPAFFAIRE')) then SetControlText('TGPAFFAIRE', 'Code Affaire');
    end;
  end;

end;

Procedure TOF_BTACCEPTAFF_MUL.ResponsableOnElipsisClick(Sender : Tobject);
begin

  Responsable.text := AFLanceFiche_Rech_Ressource('ARS_RESSOURCE='+ Responsable.text,'TYPERESSOURCE=SAL; ACTION=RECH');

end;

procedure TOF_BTACCEPTAFF_MUL.OnClose ;
begin
  Inherited ;
end ;

procedure TOF_BTACCEPTAFF_MUL.OnDisplay () ;
begin
  Inherited ;
end ;

procedure TOF_BTACCEPTAFF_MUL.OnCancel () ;
begin
  Inherited ;
end ;

procedure TOF_BTACCEPTAFF_MUL.FlisteDblClick (Sender : TObject);
var TheAffaire  : string;
    TheDateArretee : string;
    TOBAFFRAD   : TOB;
    QQ          : TQuery;
begin

  QQ := TFmul(ecran).Q;

  Try
    if ecran.Name = 'BTSAISRESTDEP_MUL' then
    begin
      TobRAD := TOB.Create('RESTADEP',nil,-1);

      TheDateArretee := QQ.findfield('RAD_DATEARRETEE').AsString;
      TheAffaire:= QQ.FindField('AFF_AFFAIRE').AsString;

      TOBAFFRAD := TOB.Create ('AFFAIRAD',TobRAD,-1);
      TOBAFFRAD.AddChampSupValeur('AFFAIRE', TheAffaire);
      TOBAFFRAD.AddChampSupValeur('DATEARRETEE', TheDateArretee);
      //
      SaisieResteADepenser;
    end
    else
    begin
      TheAffaire := QQ.FindField('AFF_AFFAIRE').AsString;
      if PgiAsk ('Désirez-vous accepter l''affaire ?')=MrYes then AcceptationAffaire(TheAffaire);
    end;
  except
    PGIBox('La date d''arrêté n''est pas paramétrée dans la liste de présentation', 'Saisie R.A.D');
    TheDateArretee := ''
  end;

  TFmul(ecran).BChercheClick(ecran);

  FreeAndNil(TOBRad);

end;

procedure TOF_BTACCEPTAFF_MUL.SaisieResteADepenser;
begin

  TheTOB := TobRAD;

  AglLanceFiche('BTP','BTSAISIERESTDEP','','',''); // Appels

  TheTOB := nil;

end;

procedure TOF_BTACCEPTAFF_MUL.AcceptationAffaire(TheAffaire : string);
begin

  //Mise à jour de l'affaire !!!
  if (ExecuteSQL ('UPDATE AFFAIRE SET AFF_ETATAFFAIRE="ACP" WHERE AFF_AFFAIRE="'+ TheAffaire +'"')< 1) then
   	 PgiBox ('Erreur Pendant la mise à jour');

end;

Procedure TOF_BTACCEPTAFF_MUL.MultiAcceptation(Sender : TOBJect);
var TheAffaire  : String;
    i           : integer;
    QQ			    : TQuery;
{$IFDEF EAGLCLIENT}
    L 			    : THGrid;
{$ELSE}
    //Ext 		  : String;
    L 			    : THDBGrid;
{$ENDIF}
    TOBAFFRAD   : TOB;
begin
  Inherited ;

  //controle si au moins un éléments sélectionné
	if (TFmul(ecran).FListe.NbSelected=0)and(not TFmul(ecran).FListe.AllSelected) then
  begin
	  PGIInfo('Aucun chantier sélectionné','');
  	exit;
  end;

  if ecran.Name <> 'BTSAISRESTDEP_MUL' then
  begin
    if PgiAsk ('Désirez-vous accepter l(les)''affaire(s) sélectionnée(s) ?') <> MrYes then
    Begin
      SourisNormale;
      TFmul(ecran).BChercheClick(ecran);
      exit;
    end;
  end;

  TobRAD := TOB.Create ('RESTADEP',nil,-1);

	L:= TFmul(ecran).FListe;

  SourisSablier;

  TRY
	if L.AllSelected then
  begin
    QQ:= TFmul(ecran).Q;
    QQ.First;
    while not QQ.EOF do
    Begin
      TheAffaire:= QQ.findfield('AFF_AFFAIRE').AsString;
      if ecran.Name = 'BTSAISRESTDEP_MUL' then
      begin
        TOBAFFRAD := TOB.Create ('AFFAIRAD',TobRAD,-1);
        TOBAFFRAD.AddChampSupValeur('AFFAIRE', TheAffaire);
      end
      else
        AcceptationAffaire(TheAffaire);
      QQ.next;
    end;
  end
  else
  begin
    for i:=0 to L.NbSelected-1 do
    begin
      L.GotoLeBookmark(i);
      TheAffaire := TFMul(TFmul(ecran)).Fliste.datasource.dataset.FindField('AFF_AFFAIRE').AsString;
      if ecran.Name = 'BTSAISRESTDEP_MUL' then
      begin
        TOBAFFRAD := TOB.Create ('AFFAIRAD',TOBRAD,-1);
        TOBAFFRAD.AddChampSupValeur('AFFAIRE', TheAffaire);
      end
      else
        AcceptationAffaire(TheAffaire);
    end;
	end;
  //
  if ecran.Name = 'BTSAISRESTDEP_MUL' then
  begin
    SaisieResteADepenser;
  end;
	FINALLY
    SourisNormale;
    L.AllSelected:=False;
    TFmul(ecran).BChercheClick(ecran);
    FreeAndNil(TOBRad);
  END;

End;

procedure TOF_BTACCEPTAFF_MUL.NomsChampsAffaire(var Aff, Aff0, Aff1, Aff2, Aff3, Aff4, Aff_, Aff0_, Aff1_, Aff2_, Aff3_, Aff4_, Tiers, Tiers_:THEdit);
begin

Aff:=THEdit(GetControl('AFF_AFFAIRE'));

// MODIF LS
Aff0 := THEdit(GetControl('AFF_AFFAIRE0'));

// --
Aff1:=THEdit(GetControl('AFF_AFFAIRE1'));
Aff2:=THEdit(GetControl('AFF_AFFAIRE2'));
Aff3:=THEdit(GetControl('AFF_AFFAIRE3'));
Aff4:=THEdit(GetControl('AFF_AVENANT'));
Tiers:=THEdit(GetControl('AFF_TIERS'));

// affaire de référence pour recherche
Aff_:=THEdit(GetControl('AFF_AFFAIREREF'));
Aff1_:=THEdit(GetControl('AFFAIREREF1'));
Aff2_:=THEdit(GetControl('AFFAIREREF2'));
Aff3_:=THEdit(GetControl('AFFAIREREF3'));
Aff4_:=THEdit(GetControl('AFFAIREREF4'));

end;

procedure TOF_BTACCEPTAFF_MUL.BDefinitiveOnClick(Sender: TObject);
var TheAffaire  : String;
    Thedate     : TDatetime;
    i           : integer;
    QQ			    : TQuery;
{$IFDEF EAGLCLIENT}
    L 			    : THGrid;
{$ELSE}
    //Ext 		  : String;
    L 			    : THDBGrid;
{$ENDIF}
begin

  //controle si au moins un éléments sélectionné
	if (TFmul(ecran).FListe.NbSelected=0)and(not TFmul(ecran).FListe.AllSelected) then
  begin
	  PGIInfo('Aucun chantier sélectionné','');
  	exit;
  end;

  if PGIAsk('Désirez-vous valider définitivement les saisies sélectionnées ?', 'confirmation')=MrNo then exit;

	L := TFmul(ecran).FListe;

  SourisSablier;

  TRY
	if L.AllSelected then
  begin
    QQ:= TFmul(ecran).Q;
    QQ.First;
    while not QQ.EOF do
    Begin
      Try
        TheAffaire:= QQ.findfield('AFF_AFFAIRE').AsString;
        Thedate   := QQ.FindField('RAD_DATEARRETEE').AsDateTime;
        if QQ.Findfield('RESTEADEP').AsString = 'Saisie' then
        begin
          ValidationRAD(TheAffaire, TheDate);
        end
        else if QQ.Findfield('RESTEADEP').AsString = 'Validé' then
        begin
          PGIError('Il n''est pas possible de valider une saisie déjà validée (' + TheAffaire + ')', 'Validation des R.A.D')
        end
        else
        begin
          PGIError('Sélectionnez des affaires dont le RAD a déjà été saisi !', 'Validation des R.A.D')
        end;
      except
        PGIBox('La date d''arrêté n''est pas paramétrée dans la liste du multi-critère', 'Saisie R.A.D');
        break;
      end;
      QQ.next;
    end;
  end
  else
  begin
    for i:=0 to L.NbSelected-1 do
    begin
      try
        L.GotoLeBookmark(i);
        TheAffaire := TFMul(TFmul(ecran)).Fliste.datasource.dataset.FindField('AFF_AFFAIRE').AsString;
        Thedate    := TFMul(TFmul(ecran)).Fliste.datasource.dataset.FindField('RAD_DATEARRETEE').AsDateTime;
        if TFMul(TFmul(ecran)).Fliste.datasource.dataset.Findfield('RESTEADEP').AsString = 'Saisie' then
        begin
          ValidationRAD(TheAffaire, TheDate);
        end
        else if TFMul(TFmul(ecran)).Fliste.datasource.dataset.Findfield('RESTEADEP').AsString = 'Validé' then
        begin
          PGIError('Il n''est pas possible de valider une saisie déjà validée (' + TheAffaire + ')', 'Validation des R.A.D')
        end
        else
        begin
          PGIError('Sélectionnez des affaires dont le RAD a déjà été saisi !', 'Validation des R.A.D')
        end;
      except
        PGIBox('La date d''arrêté n''est pas paramétrée dans la liste du multi-critère', 'Saisie R.A.D');
        break;
      end;
    end;
	end;
  //
	FINALLY
    SourisNormale;
    L.AllSelected:=False;
    TFmul(ecran).BChercheClick(ecran);
  END;

end;

procedure TOF_BTACCEPTAFF_MUL.ValidationRAD(Affaire: String; DateArrete : TDateTime);
Var StSQL   : String;
    //
    Mois    : String;
    Annee   : String;
    //
    Year    : Word;
    Month   : Word;
    Day     : Word;
Begin

  DecodeDate(DateArrete, Year, Month, Day);

  Mois   := IntToStr(Month);
  Annee  := IntToStr(Year);

  //chargement de la TOB Reste à Dépenser (cas de la modification sinon blanc)
  StSQl := 'UPDATE BTRESTEADEP SET RAD_VALIDE="X" WHERE RAD_AFFAIRE="' + Affaire + '" '+
           '  AND RAD_MOIS  ="' + Mois + '" ' +
           '  AND RAD_ANNEE ="' + Annee + '" ' ;

  if (ExecuteSQL (StSQL)< 1) then PgiBox ('Erreur Pendant la mise à jour des Restes à dépenser');

end;

Initialization
  registerclasses ( [ TOF_BTACCEPTAFF_MUL ] ) ;
end.

