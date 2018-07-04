unit UTofBTHORAIRESTD;

interface
uses    StdCtrls,
				Controls,
        Classes,
        forms,
        sysutils,
        ComCtrls,
        HCtrls,
        UTOF,
        DicoAF,
        HEnt1,
        HTB97,
        UAFO_Calendrier,
        CalendRes,
        UTob,
{$IFDEF BTP}
	  		CalcOleGenericBTP,
{$ENDIF}
{$IFDEF EAGLCLIENT}
      	MaineAGL,
        eTablette,
{$ELSE}
      	Fe_main,
        dbTables,
        db,
        Tablette,
{$ENDIF}
      	HeureUtil,
        AGLInit,
        Vierge,
        UtomCalendrierRegle ;

Type
     TOF_BTHORAIRESTD = Class (TOF)
        procedure OnArgument(stArgument : String ) ; override ;
        procedure OnUpdate ; Override;
        procedure OnClose ; Override;
        procedure OnDelete ; Override;
     private

        Erreur 			: Boolean;

        OldValue 		: Double;
        OldValueS		: Double;

        TypeHoraire	: String;
        CodeRes			: String;
        CodeSalarie	: String;
        LibNom			: String;
        Standard		: String;
        TypAction		: String;

        Calendrier 	: TAFO_Calendrier;

        procedure FormatHeure (Sender: TObject);
        procedure GereEnter   (Sender: TObject);
        procedure DetailCalendrier (Sender: TObject);
        procedure HoraireVersCalendrier;
        procedure DuplicSemaine (Sender: TObject);
        procedure ChargeStandard (Sender: TObject);
        procedure EraseHoraireJour (Sender: TObject);
        Procedure CalculDuree (Numlig : integer ; Sender: TObject);
        procedure CalculTotalHebdo;
        procedure RazCalendrierStd;
        procedure BtCalregClick (Sender: TObject);
        procedure DuplicCalendrierRegle(TypeCRef,standard,TypeHoraire,CodeRes,CodeSalarie : string);
		    procedure ControlChamp(Champ: String; Valeur: Variant);

    END ;

// Attention : Les champs de la fiches ont un tag particulier afin de les traiter en
// automatique quelque soit le jour de la semaine
// les durées ont un tag 101 à 107 et les heures début et fin de 1 à 7 .

Type
     TOF_DUPLICCALENDRIER = Class (TOF)
        procedure OnArgument(stArgument : String ) ; override ;
        procedure OnUpdate ; Override;
        procedure OnClose ; Override;
     private
        Function  TestDateDuplication (ControlTest: TCustomEdit ; TypeDate: integer): Boolean;
        procedure OnExitDate (Sender: TObject);
    END ;

Procedure AFLanceFiche_HoraireStd(Argument:string);
implementation

uses UtilPaieAffaire;

procedure TOF_BTHORAIRESTD.OnArgument(stArgument : String );
Var HORAIRE : THEDIT;
    Combo   : THValComboBox;
    CTotal  : THLabel;
    BtDet   : TToolBarButton97;
    BtDuplic: TToolBarButton97;
    BtErase : TToolBarButton97;
    BtCalreg: TToolBarButton97;
    Critere : String;
    Champ		: String;
    valeur  : String;
    col			: Integer;
    lig     : integer;
begin
Inherited;

// Récupération des paramètres
// 4 paramétres sont interprétés TYPE: 'STD' 'RES' ou 'SAL'
// CODE: code de la ressource ou salarié LIBELLE : libellé associé
// STANDARD: Code standard calendrier

  TypeHoraire := 'STD';
  CodeRes := '';
  CodeSalarie := '';
  LibNom := '';
  Standard := '';
  TypAction := 'MODIFICATION'; // MV 13-10-2000
  Critere:=(Trim(ReadTokenSt(stArgument)));

  While (Critere <>'') do
  	  BEGIN
    	if Critere<>'' then
      	 BEGIN
         if pos(':',Critere)<>0 then
         	  begin
            Champ:=ReadTokenPipe(Critere,':');
            Valeur:=critere;
            end
         else
            begin
            if pos('=',Critere)<>0 then
               begin
               Champ:=ReadTokenPipe(Critere,'=');
               Valeur:= Critere;
               end;
            end;
         ControlChamp(Champ, Valeur);
         END;
    	Critere:=(Trim(ReadTokenSt(stArgument)));
    	END;


 // Affectation d'évènements sur enter et l'exit des horaires
  for lig := 1 To 7 do
      Begin
      for Col := 1 To 5 do
          Begin
          HORAIRE := THEDIT(GetControl('ACA_HORAIRE'+ IntToStr(Lig)+IntToStr(Col)));
          HORAIRE.OnExit  := FormatHeure;
          HORAIRE.OnEnter := GereEnter;
          HORAIRE.Text := '00:00';
          End;
      End;

	BtDet   := TToolBarButton97 (GetControl('BTDETCALEN'));
	BtDet.OnClick    := DetailCalendrier;

	BtDuplic:= TToolBarButton97 (GetControl('BTDUPLICATION'));
	BtDuplic.OnClick := DuplicSemaine;

	BtCalReg:= TToolBarButton97 (GetControl('BCALREGLE'));
	BtCalreg.OnClick := BtCalregClick;

	for lig := 1 to 7 do
  	  Begin
    	BtErase:= TToolBarButton97 (GetControl('BTERASE'+IntToStr(lig)));
    	BtErase.OnClick:= EraseHoraireJour;
    	End;

	Combo := THValComboBox(GetControl('ACA_STANDCALEN'));
	if (TypeHoraire = 'STD') then Combo.OnClick := ChargeStandard;

	// Affichage spécifique si calendrier ressource ou salarié
	if (TypeHoraire <> 'STD') then
  	  Begin
      if (TypeHoraire ='RES')  then
         Begin
         SetControlText('LIBELLE','Calendrier de ' + TraduitGA('la ressource'));
         SetControlText('CODE', coderes); SetControlText('LIBNOM', LibNom);
         //NCX 23/07/01
         Codesalarie := RechercheSalarieRessource (Coderes,False);
         End
      else
         Begin
         SetControlText ('LIBELLE','Calendrier du salarié ') ;
         SetControlText ('CODE', codesalarie); SetControlText('LIBNOM', LibNom);
         //NCX 23/07/01
         End;
      if (Standard <> '') then
         Begin
         SetControltext('TACA_STANDCALEN', 'Calendrier de référence');
         SetControlEnabled ('ACA_STANDCALEN', False);
         Combo.Value:= Standard;
         End
      else
         Begin
         SetControlVisible('TACA_STANDCALEN', False);
         SetControlVisible('ACA_STANDCALEN', False );
         End;
      ChargeStandard(self);
      End
  else
      Begin
      SetControlEnabled ('BTDETCALEN', False);
      SetControlEnabled ('BCALREGLE',False);
      SetControlEnabled ('BTDUPLICATION',False);
      SetControlVisible ('LIBELLE', False); SetControlVisible ('LIBNOM', False);
      SetControlVisible ('CODE', False);
	    if (Standard <> '') then
    	    Begin
          Combo.Value:= Standard; //Combo.Enabled := False;  SetControlVisible ('BTCREATSTANDCALEN',False);
          ChargeStandard(self);
          End;
      End;

	SetControlText('TYPECODE',TypeHoraire);
	CTotal := THLabel(GetControl('ACA_HORAIRETOTAL'));
	CTotal.Caption := Format ( '%2.2f' , [0.0]);

	// MV debut 13-10-2000
  if TypAction = 'CONSULTATION' then
     begin
     //la fiche devient non saisissable
     for lig:=1 To 7 do
         for Col:=1 To 5 do
             SetControlEnabled('ACA_HORAIRE'+ IntToStr(Lig)+IntToStr(Col),TypAction='MODIFICATION');
     SetControlEnabled('ACA_STANDCALEN',TypAction='MODIFICATION');
     for lig:=1 To 7 do
         SetControlEnabled('BTERASE'+ IntToStr(Lig),TypAction='MODIFICATION');
     end;

  SetControlEnabled('CODE',(TypAction='MODIFICATION') and (TypeHoraire <> 'SAL') or ((typeHoraire='SAL')and (codesalarie='')));

  CalculTotalHebdo;

// MV fin 13-10-2000
{$ifdef PAIEGRH}   // ajout mcd 08/01/03
  SetControlVIsible ('BCALREGLE',False);
{$endif}

End;

Procedure TOF_BTHORAIRESTD.ControlChamp(Champ : String; Valeur : Variant);
Begin

  if Champ = 'TYPE' then TypeHoraire := Valeur;

  if Champ = 'CODE' then
     if TypeHoraire = 'SAL' then
        CodeSalarie := Valeur
     else
        CodeRes 		:= Valeur;

  if Champ = 'LIBELLE'  then LibNom := Valeur;

  if Champ = 'STANDARD' then Standard := Valeur;

  if Champ = 'ACTION'   then TypAction:= Valeur;

end;

procedure TOF_BTHORAIRESTD.OnUpdate;
Var Combo   : THValcomboBox;
BEGIN
Inherited;

	if Calendrier = nil then exit; //mv 16-10-2000

	if (TypeHoraire='STD') then
     SetFocusControl('ACA_STANDCALEN')
	else
     SetFocusControl('CODE');

	Combo := THValComboBox (GetControl('ACA_STANDCALEN'));

	if (((Combo.Text<>'') And (TypeHoraire='STD')) Or (TypeHoraire<>'STD')) then
     Begin
     HoraireVersCalendrier;
     if (Calendrier.TobCalenSemaine.IsOneModifie) then
     if (PGIAskAF('Voulez-vous enregistrer les modifications','Calendrier')= mrYes) then
        Calendrier.UpdateCalendrierSemaine;
    END;

	//mv pour récupérer en retour le nom du calendrier choisi. La tob a un seul champ : CALENDRIER_MAJ
	if (LaTOB <> NIL) then
  	 // PL le 24/09/02 : en Gestion interne/GA le 'CALENDRIER_MAJ' n'existe pas...
     if LaTOB.FieldExists('CALENDRIER_MAJ') then
        Latob.PutValue('CALENDRIER_MAJ',Combo.value);

END;


procedure TOF_BTHORAIRESTD.OnDelete;
Var Combo : THValComboBox;
BEGIN
Combo := THValComboBox (GetControl ('ACA_STANDCALEN') );
if (Calendrier <> Nil) then
    BEGIN
    if (PGIAskAF('Confirmer vous la suppression de tous les horaires de ce calendrier','Calendrier')= mrYes) then
        BEGIN
        If CodeRes='' then Coderes := '***';
        If CodeSalarie='' then CodeSalarie := '***';
        DeleteCalendrier (TypeHoraire,Combo.Value,CodeRes,CodeSalarie);
        Calendrier.Free; Calendrier := Nil;
        ExecuteSQL('DELETE FROM CALENDRIERREGLE WHERE ACG_STANDCALEN="'+Combo.Value+
                   '" AND ACG_RESSOURCE="***" AND ACG_SALARIE="***"');
        END;
    END;

	if (TypeHoraire<> 'STD') then
  	 TFVierge(Ecran).Close
	else
     RazCalendrierStd;

END;


procedure TOF_BTHORAIRESTD.OnClose;
BEGIN
Inherited;
	OnUpdate;
	Calendrier.Free;
  Calendrier := Nil;
END;

procedure TOF_BTHORAIRESTD.FormatHeure (Sender: TObject);
Var Tmp		: Double;
		HSav	: TDateTime;
begin

	HSav := StrToTime(THEdit(Sender).Text);
  Tmp := TimeToFloat(HSav, true);

	if (Tmp = 0) then
		 begin
  	 if (OldValueS <> Tmp) then
    		CalculDuree(THEdit(Sender).Tag, Sender );
    		exit;
  	 end;

  // MAJ du cumul global
  if (THEdit(Sender).Tag < 100) then
     if (OldValue <> Tmp) then CalculDuree(THEdit(Sender).Tag, Sender)
  else
     if (OldValue <> Tmp) then CalculTotalHebdo;

END;

procedure TOF_BTHORAIRESTD.GereEnter (Sender: TObject);
Var Combo   : THValcomboBox;
BEGIN

	Combo := THValComboBox(GetControl('ACA_STANDCALEN'));
	if (Combo.Text='') and (TypeHoraire='STD') then
     BEGIN
     PgiBoxAF('Vous devez saisir un standard de calendrier',TitreHalley);
     Combo.SetFocus;
     END;

  OldValueS := StrTimeToFloat(THEdit(Sender).Text);

	if (OldValueS <> 0) then OldValue := StrTimeToFloat(THEdit(Sender).Text);

  // Affichage au format double
  //THEdit(Sender).Text := StrFloatToTime(THEdit(Sender).Text));

	if Erreur then
     BEGIN
     OldValue := -1;
     Erreur := False;
     END;

END;

Procedure TOF_BTHORAIRESTD.DetailCalendrier (Sender: TObject);
Var Combo   : THValcomboBox;
BEGIN
Combo := THValComboBox (GetControl('ACA_STANDCALEN'));
if (TypeHoraire='STD') then SetFocusControl('ACA_STANDCALEN') else SetFocusControl ('CODE');
HoraireVersCalendrier;
CalendrierDetail(TypeHoraire,Calendrier,Combo.Value,Combo.Text,CodeRes,CodeSalarie,LibNom);
END;

Procedure  TOF_BTHORAIRESTD.CalculDuree (NumLig : integer ; Sender: TObject);
Var CDeb1 	: THEDIT;
    CFin1 	: THEDIT;
    CDeb2 	: THEDIT;
    CFin2 	: THEDIT;
    Plage1  : double;
    Plage2 	: double;
BEGIN

	Plage1 := 0;
  Plage2 := 0;

	CDeb1  := THEDIT (GetControl('ACA_HORAIRE'+ IntToStr(NumLig)+ '1'));
	CFin1  := THEDIT (GetControl('ACA_HORAIRE'+ IntToStr(NumLig)+ '2'));
	CDeb2  := THEDIT (GetControl('ACA_HORAIRE'+ IntToStr(NumLig)+ '3'));
	CFin2  := THEDIT (GetControl('ACA_HORAIRE'+ IntToStr(NumLig)+ '4'));

	if (CDeb2.Text < CFin1.Text) And (CDeb2.Text <> '') And (StrTimeToFloat(CDeb2.Text)<> 0) then
     BEGIN
     PgiBoxAF('L''horaire de début de la plage Horaire 2 est inférieur à la plage 1',TitreHalley);
     THEdit(Sender).Text := '';
     THEdit(Sender).SetFocus;
     Erreur := True;
     Exit;
     END;

	if (CFin1.Text <> '') and (CDeb1.Text <> '') then
  	 Plage1 := CalculEcartHeure(StrTimeToFloat(CDeb1.Text), StrTimeToFloat (CFin1.Text));

	if (CFin2.Text <> '') and (CDeb2.Text <> '') then
  	 Plage2 := CalculEcartHeure ( StrTimeToFloat (CDeb2.Text) , StrTimeToFloat (CFin2.Text) );

	Plage1 := AjouteHeure(Plage1, Plage2);

	SetControlText('ACA_HORAIRE'+ IntToStr(NumLig)+ '5',FloatToStrTime(Plage1,FormatHM));

	CalculTotalHebdo;

END;

procedure TOF_BTHORAIRESTD.HoraireVersCalendrier;
Var Horaire : TAFO_Horaire;
    Lig     : Integer;
    Col     : integer;
    CC      : THEdit;
BEGIN

	Horaire:= TAFO_Horaire.Create;

	for lig:=1 To 7 do
  	  Begin
    	for Col:=1 To 5 do
      	  Begin
        	CC := THEDIT(GetControl('ACA_HORAIRE'+ IntToStr(Lig)+IntToStr(Col)));
        	Case Col of        // Attention stockage en 100ème d'heures...
          	  1 : Horaire.Hdeb1  := StrTimeToFloat(CC.Text, True);
            	2 : Horaire.Hfin1  := StrTimeToFloat(CC.Text, True);
            	3 : Horaire.Hdeb2  := StrTimeToFloat(CC.Text, True);
            	4 : Horaire.Hfin2  := StrTimeToFloat(CC.Text, True);
            	5 : Horaire.Hduree := StrTimeToFloat(CC.Text, True);
          End;
      		End;
    	Calendrier.SetCalendrierSemaine(Horaire,lig);
    	End;

	Horaire.free;  //mcd 06/06/01

END;

Procedure  TOF_BTHORAIRESTD.CalculTotalHebdo;
Var SommeDuree  : Double;
    PlageHeure	: Double;
    i           : integer;
    st          : String;
    CTotal      : THLabel;
BEGIN

	SommeDuree := 0;

	for i:=1 to 7 do
  	  Begin
	    st := GetControlText('ACA_HORAIRE'+IntTostr(i)+'5');
      PlageHeure := TimeToFloat(StrToTime(st), true);
  	  SommeDuree := SommeDuree + PlageHeure;
	    End;

	//st := Format ( '%2.2f' , [BTPHeureBase100To60(SommeDuree)]);
	//st := FindEtReplace(st,' ','0',True);

	CTotal := THLabel(GetControl('ACA_HORAIRETOTAL'));
  CTotal.Caption := FloatToStr(HeureBase60To100(SommeDuree));

END;

procedure TOF_BTHORAIRESTD.ChargeStandard (Sender: TObject);
Var TD      : Tob;
    CC      : THEdit;
    Combo   : THValComboBox;
    jour		: Integer;
    lig			: Integer;
    col  		: integer;
BEGIN

  //mcd 03/04/2003 si on clic sur le chgmt std, il faure remttre à blanc le n° std
	Standard :='';

	if (Calendrier <> Nil) then
  	  BEGIN
    	HoraireVersCalendrier;
    	If (Calendrier.TobCalenSemaine.IsOneModifie) then
      	  Begin
        	If (PGIAskAF('Voulez-vous enregistrer les modifications','Calendrier')= mrYes) Then
          	  Calendrier.UpdateCalendrierSemaine;
        	End;
    	END;

	// Récupération des composants utilisés
	Combo := THValComboBox (GetControl ('ACA_STANDCALEN') );

	// chargement des horaires en fonction du calendrier
	if Calendrier<>NIL then BEGIN Calendrier.Free; Calendrier := nil; END;

	if (Combo.Text = '') and (TypeHoraire = 'STD') then
     Begin
     PgiBoxAF('Vous devez saisir un standard de calendrier',TitreHalley);
     Exit;
     end;

  if (TypeHoraire <> 'STD') And (Combo.Value <> '') then
     Begin
     // Test l'existance du calendrier
     if Not (ExisteCalendrier(TypeHoraire,Combo.Value,CodeRes,CodeSalarie)) then
        Begin
        if (PGIAskAF ('Création du calendrier. Voulez-vous repartir du calendrier standard ?','Calendrier')= mrYes) then
           begin
           DuplicCalendrier ('STD', Combo.Value, '', '', TypeHoraire, Combo.Value, CodeRes, CodeSalarie);
           if (V_PGI.IoError = oeOk) then
              DuplicCalendrierRegle ('STD', Combo.Value, TypeHoraire, CodeRes, CodeSalarie);
           end;
        End;
     End;

	SetControlEnabled('BTDETCALEN', True);
  SetControlEnabled('BCALREGLE',True);
  SetControlEnabled('BTDUPLICATION', True);

  Calendrier := TAFO_Calendrier.Create(TypeHoraire,Combo.Value,CodeRes,CodeSalarie,True);

  // Réinitialisation des zones de saisie
  for lig:=1 To 7 do
      Begin
      For Col:=1 To 5 do
          Begin
          CC := THEDIT(GetControl('ACA_HORAIRE'+ IntToStr(Lig)+IntToStr(Col)));
          CC.Text := '00:00';
          End;
      End;

  for lig:=0 to Calendrier.TobCalenSemaine.Detail.Count-1 do
      Begin
      TD:=Calendrier.TobCalenSemaine.Detail[lig];
      jour:=integer(TD.GetValue('ACA_JOUR'));
      if (TD.GetValue ('ACA_HEUREDEB1') <> 0) then
         Begin
         CC := THEDIT(GetControl('ACA_HORAIRE'+ IntToStr(Jour)+'1'));
         CC.Text:= FormatDateTime('hh:mm', FloatToTime(TD.GetValue('ACA_HEUREDEB1')));
         End;
      if (TD.GetValue ('ACA_HEUREFIN1') <> 0) then
         Begin
         CC := THEDIT(GetControl('ACA_HORAIRE'+ IntToStr(Jour)+'2'));
				 CC.Text:= FormatDateTime('hh:mm', FloatToTime(TD.GetValue('ACA_HEUREFIN1')));
         End;
      if (TD.GetValue ('ACA_HEUREDEB2') <> 0) then
         Begin
         CC := THEDIT(GetControl('ACA_HORAIRE'+ IntToStr(Jour)+'3'));
         CC.Text:= FormatDateTime('hh:mm', FloatToTime(TD.GetValue('ACA_HEUREDEB2')));
         End;
      if (TD.GetValue ('ACA_HEUREFIN2') <> 0) then
         Begin
         CC := THEDIT(GetControl('ACA_HORAIRE'+ IntToStr(Jour)+'4'));
         CC.Text:= FormatDateTime('hh:mm', FloatToTime(TD.GetValue('ACA_HEUREFIN2')));
         End;
      if (TD.GetValue ('ACA_DUREE') <> 0) then
         Begin
         CC := THEDIT(GetControl('ACA_HORAIRE'+ IntToStr(Jour)+'5'));
         CC.Text:= FormatDateTime('hh:mm', FloatToTime(TD.GetValue('ACA_DUREE'), True));
         End;
      end;

    CalculTotalHebdo;

    Calendrier.TobCalenSemaine.SetAllModifie(FALSE);

END;


procedure TOF_BTHORAIRESTD.EraseHoraireJour (Sender: TObject);
Var NumLig, i : integer;
    CC : THEdit;
BEGIN
NumLig := TToolBarButton97(Sender).tag;
for i := 1 to 5 do
    BEGIN
    CC := THEDIT (GetControl ('ACA_HORAIRE'+ IntToStr(NumLig)+ IntToStr(i)));
    CC.Text := '';
    END;
END;

procedure TOF_BTHORAIRESTD.RazCalendrierStd ;
Var NumLig, i : integer;
    CC : THEdit;
    CTotal : THLabel;
BEGIN

	//Remise à zéro des zones horaires de la fiche
	for Numlig := 1 to 7 do
  	  Begin
    	for i := 1 to 5 do
      	  Begin
        	CC := THEDIT (GetControl ('ACA_HORAIRE'+ IntToStr(NumLig)+ IntToStr(i)));
        	CC.Text := FormatDateTime('hh:mm', FloatToTime(0, True));
        	End;
      End;

	CTotal := THLabel(GetControl('ACA_HORAIRETOTAL'));
  CTotal.Caption := '00:00';

	Setcontroltext ('ACA_STANDCALEN', '');

END;

procedure TOF_BTHORAIRESTD.DuplicSemaine (Sender: TObject);
Var
    DateDebRef, DateFinRef, DateDebCible, DatefinCible : TDateTime;
BEGIN
AFLanceFiche_DuplicCalendrier;
if (TheTOB <> Nil) Then
    BEGIN
    if TheTOB.FieldExists('DATEDEBREF')   then DateDebRef:=TheTOB.GetValue('DATEDEBREF')     else DateDebRef:=iDate1900;
    if TheTOB.FieldExists('DATEFINREF')   then DateFinRef:=TheTOB.GetValue('DATEFINREF')     else DateFinRef:=iDate1900;
    if TheTOB.FieldExists('DATEDEBCIBLE') then DateDebCible:=TheTOB.GetValue('DATEDEBCIBLE') else DateDebCible:=iDate1900;
    if TheTOB.FieldExists('DATEFINCIBLE') then DatefinCible:=TheTOB.GetValue('DATEFINCIBLE') else DatefinCible:=iDate1900;

    Calendrier.DuplicationCalendrier (DateDebRef,DateFinRef,DateDebCible,DatefinCible);
    END;
    TheTOB.Free; TheTOB := Nil;
END;


procedure TOF_DUPLICCALENDRIER.OnArgument(stArgument : String );
Var CC : THEdit;
BEGIN
Inherited;
  // Intégration d'évènements sur l'exit des dates de duplication
  CC := THEDIT(GetControl('DEBUTREF'));
  CC.OnExit:=OnExitDate;
  CC := THEDIT(GetControl('FINREF'));
  CC.OnExit:=OnExitDate;
  CC := THEDIT(GetControl('DEBUTCIBLE'));
  CC.OnExit:=OnExitDate;
  CC := THEDIT(GetControl('FINCIBLE'));
  CC.OnExit:=OnExitDate;
END;

procedure TOF_DUPLICCALENDRIER.OnUpdate;
Var DebutRef, FinRef, DebutCible, FinCible : THEdit;
    TobDate : TOB;
BEGIN
DebutRef := THEdit( GetControl ('DebutRef'));
If (Not TestDateDuplication(DebutRef,1)) Then Exit;
FinRef := THEdit( GetControl ('FinRef'));
If (Not TestDateDuplication(FinRef,2)) Then Exit;
DebutCible := THEdit( GetControl ('DebutCible'));
If (Not TestDateDuplication(DebutCible,3)) Then Exit;
FinCible := THEdit( GetControl ('FinCible'));
If (Not TestDateDuplication(FinCible,4)) Then Exit;

If (PGIAskAF('Confirmez vous la duplication des horaires',TitreHalley)= mrYes) then
    Begin
    // Traitement de duplication des objets associés au Grid ...
    // Passage par TheTOB pour transférer des infos dans la fiche précédente
    TobDate := TOB.Create('Liste dates',Nil,-1);
    TobDate.AddChampSup('DateDebRef',False);
    TobDate.AddChampSup('DateFinRef',False);
    TobDate.AddChampSup('DateDebCible',False);
    TobDate.AddChampSup('DateFinCible',False);
    TobDate.PutValue('DateDebRef', StrToDate(DebutRef.Text));
    TobDate.PutValue('DateFinRef', StrToDate(FinRef.Text));
    TobDate.PutValue('DateDebCible', StrToDate(DebutCible.Text));
    TobDate.PutValue('DateFinCible', StrToDate(FinCible.Text));
    if theTob <> Nil then BEGIN TheTOB.Free; TheTOB := Nil; END;
    TheTOB := TobDate;
    End
Else Close;

END;

procedure TOF_DUPLICCALENDRIER.OnClose;
BEGIN

END;

Function TOF_DUPLICCALENDRIER.TestDateDuplication (ControlTest : TCustomEdit;TypeDate:integer):Boolean;
Var JourTest : integer;
    ControlDeb : TCustomEdit;
    LibelleJour : String;
Begin
Result:=True;
ControlDeb := nil;
//TypeDate :1 Début ref / 2:fin Ref / 3: debut Cible / 4:Fin cible

If ((TypeDate=1) or (TypeDate=3)) Then Begin JourTest:=2; LibelleJour:='Lundi'; End
Else
    Begin
    JourTest := 1; LibelleJour:='Dimanche';
    If (TypeDate=2) Then
        BEGIN
        ControlDeb := THEdit(GetControl('DebutRef'));
        END Else
        BEGIN
        ControlDeb := THEdit(GetControl('DebutCible'));
        END;
    End;
// Test Date valide
If (Not IsValidDate(ControlTest.Text)) Then
    Begin
    PGIBoxAF ('Date invalide','Gestion des calendriers');
    ControlTest.SetFocus; result:=False; Exit;
    End;

// cohérence entre la date de début et de fin
If (((TypeDate=2) or (TypeDate=4)) And
    (StrToDate(ControlDeb.Text)>StrToDate(ControlTest.Text))) Then
    Begin
    PGIBoxAF ('Date incompatible avec la date de début',TitreHalley);
    ControlTest.SetFocus; result:=False;  Exit;
    End;

// Test Date début semaine  = Lundi fin de semaine = dimanche
If (DayOfWeek(StrToDate(ControlTest.Text))<>JourTest) Then
    Begin
    PGIBoxAF ('La date saisie ne correspond pas à un '+LibelleJour,'Gestion des calendriers');
    ControlTest.SetFocus; result:=False;
    End;
End;

procedure TOF_DUPLICCALENDRIER.OnExitDate (Sender: TObject);
BEGIN

	// Attention le TAG des 4 dates va de 1 à 4 pour appeler cette fonction générique
	TestDateDuplication(THEdit(Sender),THEdit(Sender).Tag);

END;

procedure TOF_BTHORAIRESTD.BtCalregClick(Sender: TObject);
//var typeCal : string;
begin
if coderes = '' then coderes := '***';
If CodeSalarie = '' then codesalarie := '***';
If Standard = '' then Standard := THValComboBox(GetControl('ACA_STANDCALEN')).Value;
If TypeHoraire='STD' then LibNom := THValComboBox(GetControl('ACA_STANDCALEN')).text;
AFLanceFiche_RegleCalendrier(Standard+';'+CodeRes+';'+CodeSalarie,'ACTION='+TypAction+';'+
               'TYPE:'+TypeHoraire+';CODERES:'+CodeRes+';CODESALARIE:'+CodeSalarie+';LIBELLE:'+LibNom+';STANDARD:'+Standard);

end;

procedure TOF_BTHORAIRESTD.DuplicCalendrierRegle(TypeCRef,standard,TypeHoraire,CodeRes,CodeSalarie: string);
var
T, TNew,Tnewfille : TOB;
Q : TQuery;
i : integer;
rq, ressource, salarie : string;
begin
ressource := Coderes;
if ressource = '' then ressource := '***';
salarie := CodeSalarie;
If salarie = '' then salarie := '***';
   // OK, fait pour duplication calendrier. il faut tout lire
 rq := 'SELECT * FROM CALENDRIERREGLE WHERE ACG_STANDCALEN="'+Standard+'" AND ACG_RESSOURCE="***" AND ACG_SALARIE="***"';
 T := Tob.create('',nil,-1);
 TNew := Tob.create('',nil,-1);
 Q := OpenSql(rq,true);
 T.LoadDetailDB('CALENDRIERREGLE','','',Q,False);
 TNew.Dupliquer(T,true,true);
 T.free;
 Ferme(Q);
 If (TNew <> Nil)  and ((TNew.FindFirst(['ACG_RESSOURCE'],[ressource],False)= nil ) or (TNew.FindFirst(['ACG_SALARIE'],[salarie],False) = Nil)) then
  begin
   For i:= 0 to TNew.detail.count-1 do
     begin
      TNewfille := TNew.detail[i];
      TNewfille.PutValue('ACG_RESSOURCE',ressource);
      TNewfille.PutValue('ACG_SALARIE',salarie);
     end;
// PL le 06/01/03 : bogue si les lignes existent déjà
//  TNew.InsertDB(nil,False);
  TNew.InsertorupdateDB;
///////////////////////
  end;
  TNew.free;
end;
Procedure AFLanceFiche_HoraireStd(Argument:string);
begin
AGLLanceFiche ('BTP','BTHORAIRESTD','','',Argument);
end;

Initialization
   registerclasses([TOF_BTHORAIRESTD,TOF_DUPLICCALENDRIER]);
end.
