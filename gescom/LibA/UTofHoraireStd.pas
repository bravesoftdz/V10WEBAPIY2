unit UTofHoraireStd;

interface
uses    StdCtrls, Controls, Classes, forms, sysutils, ComCtrls, HCtrls, UTOF, Dicobtp,
        HEnt1, HTB97, UAFO_Calendrier, CalendRes, UTob,
{$IFDEF BTP}
	  CalcOleGenericBTP,
{$ENDIF}
{$IFDEF EAGLCLIENT}
      MaineAGL, eTablette,
{$ELSE}
      Fe_main,{$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}db, Tablette,
{$ENDIF}
      HeureUtil,AGLInit,Vierge,UtomCalendrierRegle ;

Type
     TOF_HORAIRESTD = Class (TOF)
        procedure OnArgument(stArgument : String ) ; override ;
        procedure OnUpdate ; Override;
        procedure OnClose ; Override;
        procedure OnDelete ; Override;
     private
        //Variables
        {MajDuree, }
        Erreur : Boolean;
        OldValue    : double;
        OldValueS   : String;
        TypeHoraire : String;
        CodeRes     : String;
        CodeSalarie : String;
        LibNom      : String;
        Standard    : String;
        TypAction   : String;
        Calendrier  : TAFO_Calendrier;

        //Procedure et Function
        procedure FormatHeure (Sender: TObject);
        procedure GereEnter   (Sender: TObject);
        procedure DetailCalendrier (Sender: TObject);
        procedure HoraireVersCalendrier;
        procedure DuplicSemaine (Sender: TObject);
        procedure ChargeStandard (Sender: TObject);
        procedure EraseHoraireJour (Sender: TObject);
        Procedure CalculDuree (Numlig : integer ; Sender: TObject);
        //procedure TestClic (Sender: TObject);
        procedure CalculTotalHebdo;
        procedure RazCalendrierStd;
        procedure BtCalregClick (Sender: TObject);
        procedure DuplicCalendrierRegle(TypeCRef,standard,TypeHoraire,CodeRes,CodeSalarie : string);
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

procedure TOF_HORAIRESTD.OnArgument(stArgument : String );
Var CC      : THEDIT;
    Combo   : THValComboBox;
    CTotal  : THLabel;
    BtDet,  BtDuplic, BtErase, BtCalreg : TToolBarButton97;
    //BValider : Ttoolbarbutton97;
    Critere, Champ, valeur  : String;
    col, lig{, X}             : integer;
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
          if (Champ = 'TYPE')   then TypeHoraire := Valeur;
          if (Champ = 'CODE')   then
             if (TypeHoraire = 'SAL') then CodeSalarie := Valeur else CodeRes := Valeur;
          if (Champ = 'LIBELLE') then LibNom := Valeur;
          if (Champ = 'STANDARD')then Standard := Valeur;
          if (Champ = 'ACTION')  then TypAction:= Valeur;
       END;
       Critere:=(Trim(ReadTokenSt(stArgument)));
    END;

  //BValider := Ttoolbarbutton97(getcontrol('BVALIDER'));
  //if BValider <> nil then
  //   BValider.OnClick := Validerclick;

  // Affectation d'évènements sur enter et l'exit des horaires + Initialisation des zones
  for lig:=1 To 7 do
      Begin
      for Col:=1 To 5 do
          Begin
          CC := THEDIT(GetControl('ACA_HORAIRE'+ IntToStr(Lig)+IntToStr(Col)));
          CC.OnExit:=FormatHeure;
          CC.OnEnter := GereEnter;
          CC.text := '00:00';
          End;
      BtErase:= TToolBarButton97 (GetControl('BTERASE'+IntToStr(lig)));
      BtErase.OnClick:= EraseHoraireJour;
      End;

  //Gestion des Evenements sur boutons Detail/duplication et Regle
  BtDet   := TToolBarButton97 (GetControl('BTDETCALEN'));
  BtDet.OnClick    := DetailCalendrier;

  BtDuplic:= TToolBarButton97 (GetControl('BTDUPLICATION'));
  BtDuplic.OnClick := DuplicSemaine;

  BtCalReg:= TToolBarButton97 (GetControl('BCALREGLE'));
  BtCalreg.OnClick := BtCalregClick;

  //for lig := 1 to 7 do
  //    BEGIN
  //    BtErase:= TToolBarButton97 (GetControl('BTERASE'+IntToStr(lig)));
  //    BtErase.OnClick:= EraseHoraireJour;
  //    END;

  Combo   := THValComboBox(GetControl('ACA_STANDCALEN'));
  if (TypeHoraire = 'STD') then Combo.OnClick := ChargeStandard;

// Affichage spécifique si calendrier ressource ou salarié
if (TypeHoraire <> 'STD') then
    BEGIN
//mv     SetControlVisible ('BTCREATSTANDCALEN',False);
   if (TypeHoraire ='RES')  then
        BEGIN
        SetControlText ('LIBELLE','Calendrier de ' + TraduitGA('la ressource'));
        SetControlText ('CODE', coderes); SetControlText('LIBNOM', LibNom);
        //NCX 23/07/01
        Codesalarie := RechercheSalarieRessource (Coderes,False);
        END
    else
        BEGIN
        SetControlText ('LIBELLE','Calendrier du salarié ') ;
        SetControlText ('CODE', codesalarie); SetControlText('LIBNOM', LibNom);
        //NCX 23/07/01
        //Coderes := RechercheSalarieRessource (codesalarie,True);
        END;

    if (Standard <> '') then
        BEGIN
        SetControltext('TACA_STANDCALEN', 'Calendrier de référence');
        SetControlEnabled ('ACA_STANDCALEN', False);
        Combo.Value:= Standard;
        END
    else
        BEGIN
        SetControlVisible('TACA_STANDCALEN', False);
        SetControlVisible('ACA_STANDCALEN', False );
        END;
    ChargeStandard(self);
    END
else
    BEGIN
    SetControlEnabled ('BTDETCALEN', False);
    SetControlEnabled ('BCALREGLE',False);
    SetControlEnabled ('BTDUPLICATION',False);
    SetControlVisible ('LIBELLE', False); SetControlVisible ('LIBNOM', False);
    SetControlVisible ('CODE', False);
    if (Standard <> '') then
        BEGIN
        Combo.Value:= Standard; //Combo.Enabled := False;  SetControlVisible ('BTCREATSTANDCALEN',False);
        ChargeStandard(self);
        END;
    END;

    if (Standard <> '') then CalculTotalHebdo;

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
//mv       SetControlEnabled('BTCREATSTANDCALEN',TypAction='MODIFICATION');
       for lig:=1 To 7 do
                SetControlEnabled('BTERASE'+ IntToStr(Lig),TypAction='MODIFICATION');
       end;
    SetControlEnabled('CODE',(TypAction='MODIFICATION') and (TypeHoraire <> 'SAL') or ((typeHoraire='SAL')and (codesalarie='')));

// MV fin 13-10-2000
{$ifdef PAIEGRH}   // ajout mcd 08/01/03
SetControlVIsible ('BCALREGLE',False);
{$endif}

End;

procedure TOF_HORAIRESTD.OnUpdate;
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
    BEGIN
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


procedure TOF_HORAIRESTD.OnDelete;
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
if (TypeHoraire<> 'STD') then TFVierge(Ecran).Close
else
    BEGIN
    RazCalendrierStd;
    END;

END;

procedure TOF_HORAIRESTD.OnClose;
BEGIN
Inherited;
OnUpdate;
Calendrier.Free; Calendrier := Nil;
END;

procedure TOF_HORAIRESTD.FormatHeure (Sender: TObject);
Var Tmp : Double;
begin

  // Debut MV 13-10-2000 Exit;
  if (THEdit(Sender).Text ='') then
     begin
     if (OldValueS <> THEdit(Sender).Text) then
        CalculDuree(THEdit(Sender).Tag, Sender );
        exit;
     end;
  // Fin MV 13-10-2000


  Tmp := StrTimeToFloat(THEdit(Sender).Text);
  if Not TestHeure (tmp)   then
     BEGIN
     THEdit(Sender).Text := '00:00';
     THEdit(Sender).SetFocus;
     Exit;
     END
  else
    THEdit(Sender).Text := FloatToStrTime(tmp, FormatHM);

if (THEdit(Sender).Tag < 100) then
    BEGIN
    if (OldValue <> Tmp) then CalculDuree(THEdit(Sender).Tag, Sender );
    END
else
    BEGIN
    // MAJ du cumul global
    if (OldValue <> Tmp) then CalculTotalHebdo;
    END;

END;

procedure TOF_HORAIRESTD.GereEnter (Sender: TObject);
Var Combo   : THValcomboBox;
    CC      : THEDIT;
BEGIN

  Combo := THValComboBox(GetControl('ACA_STANDCALEN'));

  if (Combo.Text='') and (TypeHoraire='STD') then
     BEGIN
     PgiBoxAF('Vous devez saisir un standard de calendrier',TitreHalley);
     Combo.SetFocus;
     END;

  CC := THEDIT(Sender);

  OldValueS := CC.Text;

  if CC.Text <> '00:00' then OldValue := StrTimeToFloat(CC.Text);

  // Affichage au format heure
  THEdit(Sender).Text := CC.Text;
  if Erreur then
     BEGIN
     OldValue := -1;
     Erreur := False;
     END;

END;
{
Procedure TOF_HORAIRESTD.CreationStdCalendrier (Sender: TObject);
Var Combo   : THValcomboBox;
    st      : string;
BEGIN
Combo := THValComboBox (GetControl('ACA_STANDCALEN'));
st :=Combo.Value;
ParamTable  ('AFTSTANDCALEN',taCreat,0,nil);
AvertirTable('AFTSTANDCALEN');    // Rechargement de la tablette
Combo.reload; Combo.Value:= st;
END;
 }
Procedure TOF_HORAIRESTD.DetailCalendrier (Sender: TObject);
Var Combo   : THValcomboBox;
BEGIN
Combo := THValComboBox (GetControl('ACA_STANDCALEN'));
if (TypeHoraire='STD') then SetFocusControl('ACA_STANDCALEN') else SetFocusControl ('CODE');
HoraireVersCalendrier;
CalendrierDetail(TypeHoraire,Calendrier,Combo.Value,Combo.Text,CodeRes,CodeSalarie,LibNom);
END;

Procedure  TOF_HORAIRESTD.CalculDuree (NumLig : integer ; Sender: TObject);
Var CDeb1, CFin1, CDeb2, CFin2 : THEDIT;
    Plage1, Plage2 : double;
BEGIN

Plage1 := 0 ;  Plage2 := 0;

CDeb1  :=   THEDIT (GetControl ('ACA_HORAIRE'+ IntToStr(NumLig)+ '1'));
CFin1  :=   THEDIT (GetControl ('ACA_HORAIRE'+ IntToStr(NumLig)+ '2'));
CDeb2  :=   THEDIT (GetControl ('ACA_HORAIRE'+ IntToStr(NumLig)+ '3'));
CFin2  :=   THEDIT (GetControl ('ACA_HORAIRE'+ IntToStr(NumLig)+ '4'));


if  (CDeb2.Text < CFin1.Text) And (CDeb2.Text <> '') And ( StrTimeToFloat(CDeb2.Text)<> 0.0) then
    BEGIN
    PgiBoxAF('L''horaire de début de la plage Horaire 2 est inférieur à la plage 1',TitreHalley);
    THEdit(Sender).Text := '';
    THEdit(Sender).SetFocus;
    Erreur := True;
    Exit;
    END;

if  (CFin1.Text <> '') and (CDeb1.Text <> '') then
    Plage1 := CalculEcartHeure ( StrTimeToFloat (CDeb1.Text) , StrTimeToFloat (CFin1.Text) );

if  (CFin2.Text <> '') and (CDeb2.Text <> '') then
    Plage2 := CalculEcartHeure ( StrTimeToFloat (CDeb2.Text) , StrTimeToFloat (CFin2.Text) );

Plage1 := AjouteHeure (Plage1, Plage2);

SetControlText('ACA_HORAIRE'+ IntToStr(NumLig)+ '5',FloatToStrTime(Plage1,FormatHM));
CalculTotalHebdo;

END;

procedure TOF_HORAIRESTD.HoraireVersCalendrier;
Var Horaire : TAFO_Horaire;
    Lig     : Integer;
    Col     : Integer;
    CC      : THEdit;
BEGIN

  Horaire:= TAFO_Horaire.Create;

  for lig:=1 To 7 do
    BEGIN
    for Col:=1 To 5 do
        BEGIN
        CC := THEDIT(GetControl('ACA_HORAIRE'+ IntToStr(Lig)+IntToStr(Col)));
        Case Col of        // Attention stockage en 100ème d'heures...
            1 : Horaire.Hdeb1:= TimeToFloat(StrToTime(CC.Text));
            2 : Horaire.Hfin1:= TimeToFloat(StrToTime(CC.Text));
            3 : Horaire.Hdeb2:= TimeToFloat(StrToTime(CC.Text));
            4 : Horaire.Hfin2:= TimeToFloat(StrToTime(CC.Text));
            5 : Horaire.Hduree:=TimeToFloat(StrToTime(CC.Text));
            END;
        END;
    Calendrier.SetCalendrierSemaine(Horaire,lig);
    END;

    Horaire.free;  //mcd 06/06/01

END;

Procedure  TOF_HORAIRESTD.CalculTotalHebdo;
Var SommeDuree  : Double;
    i           : integer;
    st          : String;
    TempoTime   : Double;
    CTotal      : THLabel;
BEGIN
SommeDuree := 0.0;
for i:=1 to 7 do
    BEGIN
    st := GetControlText('ACA_HORAIRE'+IntTostr(i)+'5');
    if st <> '' then SommeDuree :=SommeDuree + HeureBase60To100(StrTimeToFloat(st));
    END;
st := Format ( '%2.2f' , [BTPHeureBase100To60(SommeDuree)]); st := FindEtReplace(st,' ','0',True);
CTotal := THLabel(GetControl('ACA_HORAIRETOTAL')); CTotal.Caption := st;
END;

procedure TOF_HORAIRESTD.ChargeStandard (Sender: TObject);
Var TD      : Tob;
    CC      : THEdit;
    Combo   : THValComboBox;
    jour, lig, col  : integer;
BEGIN

  //mcd 03/04/2003 si on clic sur le chgmt std, il faure remttre à blanc le n° std
  Standard :='';
  if (Calendrier <> Nil) then
     Begin
     HoraireVersCalendrier;
     If (Calendrier.TobCalenSemaine.IsOneModifie) then
        Begin
        If (PGIAskAF('Voulez-vous enregistrer les modifications','Calendrier')= mrYes) Then
            Calendrier.UpdateCalendrierSemaine;
        End;
     End;

  // Récupération des composants utilisés
  Combo := THValComboBox(GetControl('ACA_STANDCALEN'));

  // chargement des horaires en fonction du calendrier
  if Calendrier<>nil then
     Begin
     Calendrier.Free;
     Calendrier := nil;
     end;

  if (Combo.value = '') and (TypeHoraire = 'STD') then
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

  Calendrier:=TAFO_Calendrier.Create(TypeHoraire,Combo.Value,CodeRes,CodeSalarie,True);

  // Réinitialisation des zones de saisie
  for lig:=1 To 7 do
      Begin
      for Col:=1 To 5 do
          Begin
          SetControlText('ACA_HORAIRE'+ IntToStr(Lig)+IntToStr(Col), '00:00');
          End;
      End;

  for lig:=0 to Calendrier.TobCalenSemaine.Detail.Count-1 do
      Begin
      TD  :=Calendrier.TobCalenSemaine.Detail[lig];
      jour:=(integer(TD.GetValue('ACA_JOUR')));
      if TD.GetValue('ACA_HEUREDEB1') <> 0 then
         SetControlText('ACA_HORAIRE'+ IntToStr(Jour)+'1', FormatDateTime(FormatHM, FloatToTime(TD.GetValue('ACA_HEUREDEB1'))));
      if TD.GetValue('ACA_HEUREFIN1') <> 0 then
         SetControlText('ACA_HORAIRE'+ IntToStr(Jour)+'2', FormatDateTime(FormatHM, FloatToTime(TD.GetValue('ACA_HEUREFIN1'))));
      if TD.GetValue('ACA_HEUREDEB2') <> 0 then
         SetControlText('ACA_HORAIRE'+ IntToStr(Jour)+'3', FormatDateTime(FormatHM, FloatToTime(TD.GetValue('ACA_HEUREDEB2'))));
      if TD.GetValue('ACA_HEUREFIN2') <> 0 then
         SetControlText('ACA_HORAIRE'+ IntToStr(Jour)+'4', FormatDateTime(FormatHM, FloatToTime(TD.GetValue('ACA_HEUREFIN2'))));
      if TD.GetValue('ACA_DUREE') <> 0 then
         SetControlText('ACA_HORAIRE'+ IntToStr(Jour)+'5', FormatDateTime(FormatHM, FloatToTime(TD.GetValue('ACA_DUREE'))));
      End;

  CalculTotalHebdo;

  Calendrier.TobCalenSemaine.SetAllModifie(FALSE);

END;

procedure TOF_HORAIRESTD.EraseHoraireJour (Sender: TObject);
Var NumLig, i : integer;   
BEGIN
  NumLig := TToolBarButton97(Sender).tag;

  for i := 1 to 5 do
      Begin
      SetControlText('ACA_HORAIRE'+ IntToStr(NumLig)+ IntToStr(i), '00:00');
      End;

END;

procedure TOF_HORAIRESTD.RazCalendrierStd ;
Var NumLig  : Integer;
    i       : Integer;
    CTotal  : THLabel;
BEGIN

  for Numlig := 1 to 7 do
      Begin
      for i := 1 to 5 do
          Begin
          SetControlText('ACA_HORAIRE'+ IntToStr(NumLig)+ IntToStr(i), '00:00');
          End;
      End;

  CTotal := THLabel(GetControl('ACA_HORAIRETOTAL'));
  CTotal.Caption := '00:00';

  Setcontroltext('ACA_STANDCALEN', '');

END;

procedure TOF_HORAIRESTD.DuplicSemaine (Sender: TObject);
Var DateDebRef  : TdateTime;
    DateFinRef  : TDateTime;
    DateDebCible: TdateTime;
    DatefinCible: TDateTime;
BEGIN

    AFLanceFiche_DuplicCalendrier;

    if (TheTOB <> Nil) Then
       Begin
       if TheTOB.FieldExists('DATEDEBREF')   then
          DateDebRef:=TheTOB.GetValue('DATEDEBREF')
       else
          DateDebRef:=iDate1900;
       if TheTOB.FieldExists('DATEFINREF')   then
          DateFinRef:=TheTOB.GetValue('DATEFINREF')
       else
          DateFinRef:=iDate1900;
       if TheTOB.FieldExists('DATEDEBCIBLE') then
          DateDebCible:=TheTOB.GetValue('DATEDEBCIBLE')
       else
          DateDebCible:=iDate1900;
       if TheTOB.FieldExists('DATEFINCIBLE') then
          DatefinCible:=TheTOB.GetValue('DATEFINCIBLE')
       else
          DatefinCible:=iDate1900;
       Calendrier.DuplicationCalendrier (DateDebRef,DateFinRef,DateDebCible,DatefinCible);
       end;

    TheTOB.Free; TheTOB := Nil;

End;

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
Var DebutRef  : THEdit;
    FinRef    : THEdit;
    DebutCible: THEdit;
    FinCible  : THEdit;
    TobDate   : TOB;
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
Else
   Close;

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


{Procedure TOF_HORAIRESTD.TestClic (Sender: TObject);
Var Calendrier : TAFO_Calendrier;
    DateDebut , DateFin : TDateTime;
    NbJ : integer;  NBH : double;
BEGIN
// Test retour nb jour heure de travail sur la période
DateDebut := EncodeDate(2000,03,01);
DateFin := EncodeDate(2000,03,31);
Calendrier := TAFO_Calendrier.Create('RES', '', 'TEST','' ,false,DateDebut,DateFin);
Calendrier.JourHeureTravail(NBJ,NBH);

END;
}
procedure TOF_HORAIRESTD.BtCalregClick(Sender: TObject);
//var typeCal : string;
begin

  if coderes = ''      then CodeRes := '***';
  If CodeSalarie = ''  then CodeSalarie := '***';
  If Standard = ''     then Standard := THValComboBox(GetControl('ACA_STANDCALEN')).Value;
  If TypeHoraire='STD' then LibNom := THValComboBox(GetControl('ACA_STANDCALEN')).text;

  AFLanceFiche_RegleCalendrier(Standard + ';' + CodeRes + ';' +
                               CodeSalarie,'ACTION=' + TypAction +
                               ';TYPE:' + TypeHoraire +
                               ';CODERES:' + CodeRes +
                               ';CODESALARIE:' + CodeSalarie +
                               ';LIBELLE:' + LibNom +
                               ';STANDARD:' + Standard);

end;

procedure TOF_HORAIRESTD.DuplicCalendrierRegle(TypeCRef,standard,TypeHoraire,CodeRes,CodeSalarie: string);
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
AGLLanceFiche ('AFF','HORAIRESTD','','',Argument);
end;

Initialization
   registerclasses([TOF_HORAIRESTD,TOF_DUPLICCALENDRIER]);
end.
