{***********UNITE*************************************************
Auteur  ...... :
Créé le ...... : 09/07/2002
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : PGMULSTAGE ()
Mots clefs ... : TOF;PGMULSTAGE
*****************************************************************
PT1 24/04/2003 V_42 JL  Développement pour CWAS
PT2 16/11/2004 V_60 JL Modifs pour eFormation + accès maj cout prévisionnel
PT3 28/09/2007 V_7  FL  Uniquement stages du cycle en cours
PT4 22/10/2007 V_7  FL  Utilisation de la fiche EM_MULCATALOGUE à la place de EM_MULSTAGE
PT5 29/10/2007 V_8  FL  Ajout d'une coche pour les sous-niveaux hiérarchiques
PT6 03/03/2008 V_8  FL  Gestion de l'appel à l'écran en mode consultation salarié et GED
}
Unit UTOFPGEMMulStage;

Interface

Uses StdCtrls,Controls,Classes,
{$IFNDEF EAGLCLIENT}
     DBGrids, HDB,Mul,Fiche,db,dbTables,FE_Main,
{$ELSE}
     UtileAGL,MaineAgl,emul,
{$ENDIF}
     forms,sysutils,ComCtrls,HCtrls,HEnt1,HMsgBox,UTOF,Hqry,HTB97,UTob,PgoutilsFormation,EntPaie,P5Def,PGOutils2 ;

Type
TOF_PGEMMULSAL = Class (TOF)
       procedure OnArgument (stArgument: String);        Override;
       procedure OnLoad ;                               Override;
   private
       TypeUtilisat : String; //PT5
       TypeSaisie   : String; //PT6
       StTemp : string;
       MillesimeEC : String;
       {$IFNDEF EAGLCLIENT}
       Liste : THDBGrid;
       {$ELSE}
       Liste : THGrid;
       {$ENDIF}
       procedure OnClickSalarieSortie(Sender: TObject);
       procedure GrilleDBLClick(Sender : TObject);
       procedure NonNominatif(Sender : TObject);
       procedure MultiInsc(Sender : TObject);
       procedure ExitEdit(Sender: TObject);
       procedure SalarieElipsisClick(Sender : TObject); //PT6
       procedure OuvreSalarie(Sender : TObject); //PT6
       procedure OuvreGed(Sender : TObject); //PT6
  end ;


  TOF_PGEMMULSTAGE = Class (TOF)
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    private
    TypeUtilisat,MillesimeEC,TypeSaisie : String;
    {$IFNDEF EAGLCLIENT}
    Liste : THDBGrid;
    {$ELSE}
    Liste : THGrid;
    {$ENDIF}
    Q_MulStage : THQuery;
    procedure GrilleDblClick(Sender : TObject);
    procedure RecupSelection(Sender : TObject);
    procedure AjouterFavoris(Sender : TObject);
  end ;



Implementation

Uses lookup,  uGedFiles;

procedure TOF_PGEMMULSAL.MultiInsc(Sender : TObject);
var i : Integer;
    Salarie : String;
//    T : Tob;
begin
if ((Liste.nbSelected) > 0) and (not Liste.AllSelected) then
  begin
//    PGTobInscSalForm := Tob.Create('RecupInsc',Nil,-1);
//    InitMoveProgressForm(nil, 'Début du traitement', 'Veuillez patienter SVP ...', Liste.nbSelected, FALSE, TRUE);
//    InitMove(Liste.nbSelected, '');
    for i := 0 to Liste.NbSelected - 1 do
    begin
      {$IFDEF EAGLCLIENT}
//      if (TFMul(Ecran).bSelectAll.Down) then TFMul(Ecran).Fetchlestous;
      {$ENDIF}
      Liste.GotoLeBOOKMARK(i);
      {$IFDEF EAGLCLIENT}
      TFmul(Ecran).Q.TQ.Seek(Liste.Row - 1);
      {$ENDIF}
      Salarie := TFmul(Ecran).Q.FindField('PSA_SALARIE').asstring;
  //    T := Tob.Create('UnInsc',PGTobInscSalForm,-1);
//      T.AddChampSupValeur('SALARIE',Salarie);
      If TFMul(Ecran).Retour ='' then TFMul(Ecran).Retour := Salarie
      else TFMul(Ecran).Retour := TFMul(Ecran).Retour+';'+Salarie;
    end;
  end
  else if liste.AllSelected then
  begin
     TFmul(Ecran).Q.First;
    while not TFmul(Ecran).Q.EOF do
    begin
         {$IFDEF EAGLCLIENT}
         if (TFMul(Ecran).bSelectAll.Down) then TFMul(Ecran).Fetchlestous;
         {$ENDIF}
         Salarie := TFmul(Ecran).Q.FindField('PSA_SALARIE').asstring;
         If TFMul(Ecran).Retour ='' then TFMul(Ecran).Retour := Salarie
         else TFMul(Ecran).Retour := TFMul(Ecran).Retour+';'+Salarie;
         TFmul(Ecran).Q.Next;
    end;
  end;
  TFMul(Ecran).bSelectAll.Down := False;
  Liste.AllSelected := False;
  Liste.ClearSelected;
  TFMul(Ecran).BAnnulerClick(TFMul(Ecran).BAnnuler);
end;

procedure TOF_PGEMMULSTAGE.OnLoad ;
var SQL,StWhere : String;
begin
  Inherited ;
        StWhere := '';
        SetControlText('PST_MILLESIME','');
        SQL:= 'PST_ACTIF="X" AND ((PST_MILLESIME="'+MillesimeEC+'") OR (PST_MILLESIME="0000" AND PST_CODESTAGE NOT IN (SELECT PST_CODESTAGE FROM STAGE WHERE PST_MILLESIME="'+MillesimeEC+'")))';
        StWhere := SQL;
        If GetCheckBoxState('FAVORIS') = CbChecked then
        begin
             If StWhere <> '' then StWhere := StWhere+ ' AND PST_CODESTAGE IN (SELECT PCC_CODESTAGE FROM CURSUSSTAGE WHERE PCC_CURSUS="---" AND PCC_LIBELLE="'+V_PGI.UserSalarie+'")'
             else StWhere := 'PST_CODESTAGE IN (SELECT PCC_CODESTAGE FROM CURSUSSTAGE WHERE PCC_CURSUS="---" AND PCC_LIBELLE="'+V_PGI.UserSalarie+'")';
             SetControlEnabled('BFAVORIS', False);
        end
        Else
        begin
             If StWhere <> '' then StWhere := StWhere+ ' AND PST_CODESTAGE NOT IN (SELECT PCC_CODESTAGE FROM CURSUSSTAGE WHERE PCC_CURSUS="---" AND PCC_LIBELLE="'+V_PGI.UserSalarie+'")'
             else StWhere := 'PST_CODESTAGE NOT IN (SELECT PCC_CODESTAGE FROM CURSUSSTAGE WHERE PCC_CURSUS="---" AND PCC_LIBELLE="'+V_PGI.UserSalarie+'")';
             SetControlEnabled('BFAVORIS', True);
        end;
        SetControlText('XX_WHERE',StWhere);
end ;

procedure TOF_PGEMMULSTAGE.OnArgument (S : String ) ;
Var Num : Integer;
    Bt : TToolBarButton97;
begin
Inherited ;
        TFMul(Ecran).Retour := '';
        TypeSaisie := Trim(ReadTokenSt(S));
        millesimeEC := '';
        {$IFDEF EFORMATION}
        If ExisteSQL('SELECT PSE_SALARIE FROM DEPORTSAL WHERE PSE_RESPONSFOR="'+V_PGI.UserSalarie+'"') then TypeUtilisat := 'R'
        else TypeUtilisat := 'S';
        {$ENDIF}
        MillesimeEC := RendMillesimeEManager;
        If ExisteSQL('SELECT PCC_CODESTAGE FROM CURSUSSTAGE WHERE PCC_MILLESIME="'+MillesimeEC+'" AND PCC_CURSUS="---" AND PCC_LIBELLE="'+V_PGI.UserSalarie+'"') then SetControlChecked('FAVORIS',True);
        Q_MulStage := THQuery(Ecran.FindComponent('Q'));
        TFMul(Ecran).Caption := 'Catalogue des formations';
        {$IFNDEF EAGLCLIENT}
        Liste := THDBGrid(GetControl('FLISTE'));
        {$ELSE}
        Liste := THGrid(GetControl('FLISTE'));
        {$ENDIF}
        For Num  :=  1 to VH_Paie.NBFormationLibre do
        begin
                if Num >8 then Break;
                VisibiliteChampFormation (IntToStr(Num),GetControl ('PST_FORMATION'+IntToStr(Num)),GetControl ('TPST_FORMATION'+IntToStr(Num)));
        end;


        {$IFDEF EFORMATION}
        SetControlVisible('PAvance',False);
        SetControlProperty('PAvance','TabVisible',False);
        SetControlEnabled('CHOIXSTAGE',False);
        {$ENDIF}
        If TypeSaisie = 'SAISIEMASSE' then
        begin
             TFMul(Ecran).Caption := 'Sélection des stages pour le salarié';
             {$IFDEF EAGLCLIENT}
             Liste.MultiSelect := true;
             {$ENDIF}
             SetControlVisible('BSelectAll',True);
             SetControlVisible('BNONNOM',False);
             if Liste  <>  NIL then Liste.OnDblClick  :=  GrilleDblClick; //PT4
             SetControlProperty('BOuvrir','Hint','Inscrire aux stages');
             //PT4 - Début
             Bt := TToolBarButton97(GetControl('BOUVRIR'));
             If Bt <> Nil then Bt.OnClick := RecupSelection;
             //PT4 - Fin
        end
        else
        begin
              Bt := TToolBarButton97(GetControl('BFAVORIS'));
              If Bt <> Nil then
              begin
                   Bt.OnClick := AjouterFavoris;
                   Bt.Visible := True;
              end;
              if Liste  <>  NIL then Liste.OnDblClick  :=  GrilleDblClick;
        end;
       UpdateCaption(TFMul(Ecran)) ;
       SetControlText('PST_FORMATION3',CYCLE_EN_COURS_EM);//PT3
       SetControlProperty('PST_CODESTAGE', 'Plus', 'AND PST_FORMATION3="'+CYCLE_EN_COURS_EM+'"'); //PT3
end ;

procedure TOF_PGEMMULSTAGE.GrilleDblClick(Sender : TObject);
var Stage,Millesime : String;
    {$IFNDEF EFORMATION}
    Bt  :  TToolbarButton97;
    {$ENDIF}
begin
        {$IFDEF EAGLCLIENT}
        TFmul(Ecran).Q.TQ.Seek(Liste.Row-1) ;  //PT1
        {$ENDIF}
        If Q_MulStage.Eof then
        begin
                PGIBox('Vous devez sélectionner une formation',Ecran.Caption);
                Exit;
        end;
        Stage := Q_MulStage.FindField ('PST_CODESTAGE').AsString;
        Millesime := '0000';
        If TypeSaisie = 'CONSULTCAT' then AglLanceFiche('PAY','EM_CATALOGUE','',Stage+';'+Millesime,'CONSULTCAT')
        //PT4 - Début
        Else If TypeSaisie = 'SAISIEMASSE' Then
        Begin
          TFMul(Ecran).Retour := Stage;
          TFMul(Ecran).BAnnulerClick(TFMul(Ecran).BAnnuler);
        End
        //PT4 - Fin
        Else
        begin
          AglLanceFiche('PAY','EM_SAISIESTAGE','','','CWASINSCBUDGET;;'+Stage+';'+Millesime)
        end;
end;

procedure TOF_PGEMMULSTAGE.AjouterFavoris(Sender : TObject);
Var Q : TQuery;
    TobCursus : Tob;
    Stage : String;
    i,Imax : Integer;
begin
     if ((Liste.nbSelected) > 0) and (not Liste.AllSelected) then
  begin
    for i := 0 to Liste.NbSelected - 1 do
    begin
      Liste.GotoLeBOOKMARK(i);
      {$IFDEF EAGLCLIENT}
      TFmul(Ecran).Q.TQ.Seek(Liste.Row - 1);
      {$ENDIF}
      Stage := TFmul(Ecran).Q.FindField('PST_CODESTAGE').asstring;
      If Not existeSQL('SELECT PCC_CODESTAGE FROM CURSUSSTAGE WHERE PCC_CURSUS="---" AND PCC_CODESTAGE="'+Stage+'" AND PCC_LIBELLE="'+V_PGI.UserSalarie+'" AND PCC_MILLESIME="'+MillesimeEC+'"') then
      begin
          Q := OpenSQL('SELECT MAX(PCC_RANGCURSUS) FROM CURSUSSTAGE WHERE PCC_CURSUS="---" AND PCC_CODESTAGE="'+Stage+'"',True);
          if Not Q.EOF then
          begin
              IMax := Q.Fields[0].AsInteger;
              if IMax <> 0 then IMax := IMax + 1
              else IMax := 1;
          end
          else IMax := 1 ;
          Ferme(Q) ;
          TobCursus := Tob.Create('CURSUSSTAGE',Nil,-1);
          TobCursus.PutValue('PCC_CURSUS','---');
          TobCursus.PutValue('PCC_LIBELLE',V_PGI.UserSalarie);
          TobCursus.PutValue('PCC_ORDRE',-1);
          TobCursus.PutValue('PCC_RANGCURSUS',Imax);
          TobCursus.PutValue('PCC_CODESTAGE',Stage);
          TobCursus.PutValue('PCC_MILLESIME',MillesimeEC);
          TobCursus.InsertDB(Nil);
          TobCursus.Free;
      end;
    end;
  end
  else if liste.AllSelected then
  begin
     TFmul(Ecran).Q.First;
    while not TFmul(Ecran).Q.EOF do
    begin
         {$IFDEF EAGLCLIENT}
         if (TFMul(Ecran).bSelectAll.Down) then TFMul(Ecran).Fetchlestous;
         {$ENDIF}
         Stage := TFmul(Ecran).Q.FindField('PST_CODESTAGE').asstring;
         If Not existeSQL('SELECT PCC_CODESTAGE FROM CURSUSSTAGE WHERE PCC_CURSUS="---" AND PCC_CODESTAGE="'+Stage+'" AND PCC_LIBELLE="'+V_PGI.UserSalarie+'" AND PCC_MILLESIME="'+MillesimeEC+'"') then
         begin
          Q := OpenSQL('SELECT MAX(PCC_ORDRE) FROM CURSUSSTAGE WHERE PCC_CURSUS="---" AND PCC_CODESTAGE="'+Stage+'"',True);
          if Not Q.EOF then
          begin
              IMax := Q.Fields[0].AsInteger;
              if IMax <> 0 then IMax := IMax + 1
              else IMax := 1;
          end
          else IMax := 1 ;
          Ferme(Q) ;
          TobCursus := Tob.Create('CURSUSSTAGE',Nil,-1);
          TobCursus.PutValue('PCC_CURSUS','---');
          TobCursus.PutValue('PCC_LIBELLE',V_PGI.UserSalarie);
          TobCursus.PutValue('PCC_ORDRE',-1);
          TobCursus.PutValue('PCC_RANGCURSUS',Imax);
          TobCursus.PutValue('PCC_CODESTAGE',Stage);
          TobCursus.PutValue('PCC_MILLESIME',MillesimeEC);
          TobCursus.InsertDB(Nil);
          TobCursus.Free;
         end;
         TFmul(Ecran).Q.Next;
    end;
  end;
  TFMul(Ecran).bSelectAll.Down := False;
  Liste.AllSelected := False;
  Liste.ClearSelected;
  TFMul(Ecran).BChercheClick(TFMul(Ecran).BCherche);
end;

procedure TOF_PGEMMULSTAGE.RecupSelection(Sender : TObject);
var i : Integer;
    Stage : String;
begin
if ((Liste.nbSelected) > 0) and (not Liste.AllSelected) then
  begin
    for i := 0 to Liste.NbSelected - 1 do
    begin
      Liste.GotoLeBOOKMARK(i);
      {$IFDEF EAGLCLIENT}
      TFmul(Ecran).Q.TQ.Seek(Liste.Row - 1);
      {$ENDIF}
      Stage := TFmul(Ecran).Q.FindField('PST_CODESTAGE').asstring;
      If TFMul(Ecran).Retour ='' then TFMul(Ecran).Retour := Stage
      else TFMul(Ecran).Retour := TFMul(Ecran).Retour+';'+Stage;
    end;
  end
  else if liste.AllSelected then
  begin
     TFmul(Ecran).Q.First;
    while not TFmul(Ecran).Q.EOF do
    begin
         {$IFDEF EAGLCLIENT}
         if (TFMul(Ecran).bSelectAll.Down) then TFMul(Ecran).Fetchlestous;
         {$ENDIF}
         Stage := TFmul(Ecran).Q.FindField('PST_CODESTAGE').asstring;
         If TFMul(Ecran).Retour ='' then TFMul(Ecran).Retour := Stage
         else TFMul(Ecran).Retour := TFMul(Ecran).Retour+';'+Stage;
         TFmul(Ecran).Q.Next;
    end;
  end;
  TFMul(Ecran).bSelectAll.Down := False;
  Liste.AllSelected := False;
  Liste.ClearSelected;
  TFMul(Ecran).BAnnulerClick(TFMul(Ecran).BAnnuler);
end;


{TOF_PGEMMULSAL}
procedure TOF_PGEMMULSAL.GrilleDBLClick(Sender : TObject);
var Salarie,Aux : String;
Q : TQuery;
begin
	{$IFDEF EAGLCLIENT}
	TFmul(Ecran).Q.TQ.Seek(Liste.Row-1) ;
	{$ENDIF}
	Salarie := TFmul(Ecran).Q.FindField ('PSA_SALARIE').AsString;

	//PT6 - Début
	If (TypeSaisie = 'SALARIES') Then
		AglLanceFiche('PAY', 'SALARIE_PRIM', '', Salarie, 'ACTION=CONSULTATION')
	Else If (TypeSaisie = 'GED') Then
	Begin
		// Recherche de l'auxiliaire (référence du salarié pour la GED)
		Q := OpenSQL ('SELECT PSA_AUXILIAIRE FROM SALARIES WHERE PSA_SALARIE="'+Salarie+'"', True);
		If Not Q.EOF Then Aux := Q.FindField('PSA_AUXILIAIRE').AsString
		Else Aux := '';
		Ferme(Q);
		
		// Initialisation de la GED avant l'appel de la fenêtre
		InitializeGedFiles('', nil, gfmData, '', True);
		
		AglLanceFiche('PAY', 'RECHGEDPAIE', '', '', 'SALARIE='+Aux+';ACTION=CONSULTATION');
		
		// Terminaison de la GED
		FinalizeGedFiles();
	End
	Else If (TypeSaisie = 'SAISIEMASSE') Then
		AglLanceFiche('PAY','EM_SAISIESAL','','','CWASINSCBUDGET;;;'+MillesimeEC+';'+Salarie);
end;

procedure TOF_PGEMMULSAL.NonNominatif(Sender : TObject);
begin
     AglLanceFiche('PAY','EM_SAISIESALNM','','','CWASINSCBUDGET;;;'+MillesimeEC+';');
end;

procedure TOF_PGEMMULSAL.OnArgument (stArgument: String);
var
Defaut: THEdit;
Check : TCheckBox;
Num : Integer;
Bt  :  TToolbarButton97;
//TypeSaisie : String; //PT6
begin
	TypeSaisie := stArgument;
	MillesimeEC := RendMillesimeEManager;

	{$IFDEF EFORMATION}
	If ExisteSQL('SELECT PSE_SALARIE FROM DEPORTSAL WHERE PSE_RESPONSFOR="'+V_PGI.UserSalarie+'"') then TypeUtilisat := 'R'
	else TypeUtilisat := 'S';
	{$ENDIF}

	//SetControlProperty('PSA_LIBELLEEMPLOI','Plus', ' AND CC_CODE IN (SELECT PSA_LIBELLEEMPLOI FROM SALARIES LEFT JOIN DEPORTSAL ON PSA_SALARIE=PSE_SALARIE '+
	//  'WHERE (PSE_RESPONSFOR="'+V_PGI.UserSalarie+'" OR PSE_RESPONSFOR IN (SELECT PGS_RESPONSFOR FROM SERVICES LEFT JOIN SERVICEORDRE ON PGS_CODESERVICE=PSO_CODESERVICE'+
	//  ' WHERE PSO_SERVICESUP IN (SELECT PGS_CODESERVICE FROM SERVICES WHERE PGS_RESPONSFOR="'+V_PGI.UserSalarie+'"))))');
	Bt  :=  TToolbarButton97(GetControl('BNONNOM'));
	If Bt <> nil then Bt.OnClick := NonNominatif;
	SetControlText('DATEARRET',DateToStr(V_PGI.DateEntree));
	Check:=TCheckBox(GetControl('CKSORTIE'));
	Check.OnClick:=OnClickSalarieSortie;
	For Num := 1 to VH_Paie.PGNbreStatOrg do
	begin
		if Num >4 then Break;
		VisibiliteChampSalarie (IntToStr(Num),GetControl ('PSA_TRAVAILN'+IntToStr(Num)),GetControl ('TPSA_TRAVAILN'+IntToStr(Num)));
	end;
	VisibiliteStat (GetControl ('PSA_CODESTAT'),GetControl ('TPSA_CODESTAT')) ; //PT3
	Defaut:=ThEdit(getcontrol('PSA_SALARIE'));
	If Defaut<>nil then Defaut.OnExit:=ExitEdit;
	{$IFNDEF EAGLCLIENT}
	Liste := THDBGrid(GetControl('FLISTE'));
	{$ELSE}
	Liste := THGrid(GetControl('FLISTE'));
	{$ENDIF}

	//PT6 - Début
	If (TypeSaisie = 'SALARIES') Or (TypeSaisie = 'GED') then
	Begin
//		SetControlVisible('SOUSNIVEAUX', False);
		SetControlVisible('BNONNOM',     False);

		Defaut := THEdit (GetControl('PSA_SALARIE'));
		If Defaut <> Nil Then Defaut.OnElipsisClick := SalarieElipsisClick;

		// Attention : Type Responsable forcé!
		SetControlProperty('PSA_LIBELLEEMPLOI','Plus', ' AND CC_CODE IN (SELECT PSA_LIBELLEEMPLOI FROM SALARIES LEFT JOIN DEPORTSAL ON PSA_SALARIE=PSE_SALARIE '+
		'WHERE (PSA_DATESORTIE IS NULL OR PSA_DATESORTIE>="'+UsdateTime(V_PGI.DateEntree)+'" OR PSA_DATESORTIE<="'+UsDateTime(IDate1900)+'") AND ('+
		AdaptByRespEmanager ('R','PSE',V_PGI.UserSalarie,'ABS',False)+ ' OR ' +
		AdaptByRespEmanager ('R','PSE',V_PGI.UserSalarie,'FOR',False)+ ' OR ' +
		AdaptByRespEmanager ('R','PSE',V_PGI.UserSalarie,'VAR',False)+ ' OR ' +
		'))');

		Bt  :=  TToolbarButton97(GetControl('BOuvrir'));
		If (TypeSaisie = 'SALARIES') Then
		Begin
			If Bt <> Nil then Bt.OnClick := OuvreSalarie;
		End
		Else If (TypeSaisie = 'GED') Then
		Begin
			If Bt <> Nil then Bt.OnClick := OuvreGed;
		End;
	End;
	//PT6 - Fin
	
	If TypeSaisie = 'SAISIEMASSE' then
	begin
		SetControlText('XX_WHERERESP','PSA_SALARIE IN (SELECT PSE_SALARIE FROM DEPORTSAL WHERE PSE_RESPONSFOR="'+V_PGI.UserSalarie+'")');
		TFMul(Ecran).Caption := 'Sélection des salariés pour la formation';
		{$IFDEF EAGLCLIENT}
		Liste.MultiSelect := true;
		{$ENDIF}
		SetControlVisible('BSelectAll',True);
		SetControlVisible('BNONNOM',False);
		Bt  :=  TToolbarButton97(GetControl('BOuvrir'));
		If Bt <> Nil then Bt.OnClick := MultiInsc;
		SetControlProperty('BOuvrir','Hint','Inscrire les salariés');
	end
	else
	begin
		if Liste  <>  NIL then Liste.OnDblClick  :=  GrilleDblClick;
	end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : FLO
Créé le ...... : 03/03/2008 / PT6
Modifié le ... :   /  /
Description .. : Clic sur l'elipsis Salariés
Mots clefs ... :
*****************************************************************}
procedure TOF_PGEMMULSAL.SalarieElipsisClick(Sender : TObject);
var
  StFrom, StWhere: string;
begin
  StWhere := '(PSA_DATESORTIE<="' + UsDateTime(IDate1900) + '" OR PSA_DATESORTIE IS NULL OR PSA_DATESORTIE>="' + UsDateTime(V_PGI.DateEntree) + '")';
  StFrom := 'SALARIES LEFT JOIN DEPORTSAL ON PSA_SALARIE=PSE_SALARIE';
  // Attention : Type Responsable forcé
  StWhere := StWhere + ' AND (' + AdaptByRespEmanager ('R','PSE',V_PGI.UserSalarie,'ABS',False) + ' OR ' +
  			 AdaptByRespEmanager ('R','PSE',V_PGI.UserSalarie,'FOR',False) + ' OR ' + 
  			 AdaptByRespEmanager ('R','PSE',V_PGI.UserSalarie,'VAR',False) + ')';
  LookupList(THEdit(Sender), 'Liste des salariés', StFrom, 'PSA_SALARIE', 'PSA_LIBELLE,PSA_PRENOM', StWhere, 'PSA_SALARIE', TRUE, -1);
end;

procedure TOF_PGEMMULSAL.ExitEdit(Sender: TObject);
var edit : thedit;
begin
edit:=THEdit(Sender);
if edit <> nil then	//AffectDefautCode que si gestion du code salarié en Numérique
    if (VH_Paie.PgTypeNumSal='NUM') and (length(Edit.text)<11) and (isnumeric(edit.text)) then
    edit.text:=AffectDefautCode(edit,10);
end;



procedure TOF_PGEMMULSAL.OnLoad ;
var DateArret : TdateTime;
    StDateArret : string;
    {$IFDEF EMANAGER}
    Whereresp,Where : String;
    MultiNiveau : Boolean;
    {$ENDIF}
begin
   SetControlText('XX_WHERE',StTemp);
   if  TCheckBox(GetControl('CKSORTIE'))<>nil then
   Begin
     {PT5 mise en commentaire}
     if (GetControlText('CKSORTIE')='X') and (IsValidDate(GetControlText('DATEARRET')))then   //DEB PT2
     Begin
          DateArret:=StrtoDate(GetControlText('DATEARRET'));
          StDateArret:=' AND (PSA_DATESORTIE>="'+UsDateTime(DateArret)+'" OR PSA_DATESORTIE="'+UsdateTime(Idate1900)+'" OR PSA_DATESORTIE IS NULL) ';
          // PT5 13/05/2002 PH V575 Ajout controle date entrée par rapport à la date d'arreté
          StDateArret:=StDateArret + ' AND PSA_DATEENTREE <="'+UsDateTime(DateArret)+'"';
          SetControlText('XX_WHERE',StTemp+StDateArret);
     End                                             //FIN PT2
     else
          SetControlText('XX_WHERE',StTemp);
   End
   Else
     StDateArret:='';

   //PT5 - Début
   //PT6 - Début
{ggg
   If (TypeSaisie = 'SALARIES') Or (TypeSaisie = 'GED') Then
   Begin
    SetControlText('XX_WHERERESP','PSA_SALARIE IN (SELECT PSE_SALARIE FROM DEPORTSAL WHERE PSE_RESPONSFOR="'+V_PGI.UserSalarie+'"'+
                    'OR PSE_RESPONSABS="'+V_PGI.UserSalarie+'" OR PSE_RESPONSVAR="'+V_PGI.UserSalarie+'")');
   End
   //PT6 - Fin
   Else
   Begin
}
   If GetControlText('SOUSNIVEAUX') = 'X' Then
   Begin
     If TypeUtilisat = 'R' then
      SetControlText('XX_WHERERESP','PSA_SALARIE IN (SELECT PSE_SALARIE FROM DEPORTSAL WHERE PSE_RESPONSFOR="'+V_PGI.UserSalarie+'" OR PSE_RESPONSFOR IN '+
      '(SELECT PGS_RESPONSFOR FROM SERVICES WHERE PGS_CODESERVICE IN '+
      '(SELECT PSO_CODESERVICE FROM SERVICES LEFT JOIN SERVICEORDRE ON PGS_CODESERVICE=PSO_SERVICESUP AND'+
      ' PGS_RESPONSFOR="'+V_PGI.UserSalarie+'")))')
     Else
      SetControlText('XX_WHERERESP','PSA_SALARIE IN (SELECT PSE_SALARIE FROM DEPORTSAL WHERE PSE_RESPONSFOR IN '+
      '(SELECT PGS_RESPONSFOR FROM SERVICES LEFT JOIN SERVICEORDRE ON PGS_CODESERVICE=PSO_CODESERVICE '+
      'WHERE PGS_SECRETAIREFOR="'+V_PGI.UserSalarie+'" OR PSO_SERVICESUP IN (SELECT PGS_CODESERVICE FROM SERVICES WHERE PGS_SECRETAIREFOR="'+V_PGI.UserSalarie+'")))');
   End
   Else
   Begin
     If TypeUtilisat = 'R' then
      SetControlText('XX_WHERERESP','PSA_SALARIE IN (SELECT PSE_SALARIE FROM DEPORTSAL WHERE PSE_RESPONSFOR="'+V_PGI.UserSalarie+'")')
     Else
      SetControlText('XX_WHERERESP','PSA_SALARIE IN (SELECT PSE_SALARIE FROM DEPORTSAL WHERE PSE_RESPONSFOR IN (SELECT PGS_RESPONSFOR FROM SERVICES WHERE PGS_SECRETAIREFOR="'+V_PGI.UserSalarie+'"))');
   End;
{ggg
   End;
}   
   //PT5 - Fin
end;

procedure TOF_PGEMMULSAL.OnClickSalarieSortie(Sender: TObject);
begin
SetControlenabled('DATEARRET',(GetControltext('CKSORTIE')='X'));
SetControlenabled('TDATEARRET',(GetControltext('CKSORTIE')='X'));
end;

{***********A.G.L.***********************************************
Auteur  ...... : FLO
Créé le ...... : 03/03/2008 / PT6
Modifié le ... :   /  /
Description .. : Ouverture de la fiche Salarié
Mots clefs ... :
*****************************************************************}
procedure TOF_PGEMMULSAL.OuvreSalarie(Sender : TObject);
var Salarie : String;
begin
	{$IFDEF EAGLCLIENT}
	TFmul(Ecran).Q.TQ.Seek(Liste.Row - 1);
	{$ENDIF}
	Salarie := TFmul(Ecran).Q.FindField('PSA_SALARIE').asstring;
	AglLanceFiche('PAY', 'SALARIE_PRIM', '', Salarie, 'ACTION=CONSULTATION')
end;

{***********A.G.L.***********************************************
Auteur  ...... : FLO
Créé le ...... : 03/03/2008 / PT6
Modifié le ... :   /  /
Description .. : Chargement de la GED du salarié
Mots clefs ... :
*****************************************************************}
procedure TOF_PGEMMULSAL.OuvreGed(Sender : TObject);
var Salarie, Aux : String;
Q : TQuery;
begin
      {$IFDEF EAGLCLIENT}
      TFmul(Ecran).Q.TQ.Seek(Liste.Row - 1);
      {$ENDIF}
      Salarie := TFmul(Ecran).Q.FindField('PSA_SALARIE').asstring;

      Q := OpenSQL ('SELECT PSA_AUXILIAIRE FROM SALARIES WHERE PSA_SALARIE="'+Salarie+'"', True);
      If Not Q.EOF Then Aux := Q.FindField('PSA_AUXILIAIRE').AsString
      Else Aux := '';
      Ferme(Q);

	  // Initialisation de la GED avant l'appel de la fenêtre
      InitializeGedFiles('', nil, gfmData, '', True);

      If Aux <> '' Then AglLanceFiche('PAY', 'RECHGEDPAIE', '', '', 'SALARIE='+Aux+';ACTION=CONSULTATION');

	  // Terminaison de la GED
  	  FinalizeGedFiles();
End;

Initialization
  registerclasses ( [ TOF_PGEMMULSAL,TOF_PGEMMULSTAGE ] ) ;
end.
