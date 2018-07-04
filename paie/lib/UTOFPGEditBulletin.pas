{***********UNITE*************************************************
Auteur  ...... : PAIE PGI
Créé le ...... : 11/07/2001
Modifié le ... : //
Description .. :
Mots clefs ... : PAIE;BULLETIN
******************************************************************}
{
PT1   : 11/07/2001 SB V540 Suppression de la TOF : TOF_PGEDITBULSAL
                           Jusqu'à présent on utilisait 2 fiches QRS1 pour
                           éditer le bulletin
                           - une à partir du menu
                           - l'autre à partir de la saisie du bulletin..
                           Nous utiliserons qu'une seule fiche
PT2   : 31/07/2001 SB V547 Modification champ suffixe MODEREGLE
PT3   : 06/12/2001 SB V563 En saisie de bulletin ts les bulletins de la session
                           sont édités
PT4   : 07/12/2001 SB V563 Fiche de bug n°355 Les montants additionnés =0 ne
                           s'imprime pas
PT5   : 28/12/2001 SB V571 Edition des rubriques alimentant qu'une base
PT6   : 08/01/2002 SB V571 Edition du bulletin compl en saisie de bulletin
PT7-1 : 28/01/2002 SB V571 Fiche de bug n°244 Edition périodique
PT7-2 : 28/01/2002 SB V571 Fiche de bug n°244 Sélection alphanumérique du code
                           salarié
PT8   : 26/03/2002 SB V571 Fiche de bug n°422 Edition de la civilité
PT9   : 22/10/2002 VG V585 Version S3 - FQ N° 10286
PT10  : 20/11/2002 SB V591 Traitements Etats chainés
PT11  : 18/02/2003 SB V_42C FQ 10071 Modification requête édition
PT12-1: 19/03/2003 SB V595  FQ 10544 Ajout de la définition juridique de l'établissement
PT12-2: 19/03/2003 SB V595  FQ 10570 Ajout coche pour exclure les SLD
PT13    09/05/2003 SB V595  FQ 10645 Ajout lib organisme sur rub de cotisation
PT14    23/06/2003 SB V_42  FQ 10713 Ajout champ denominateur et numerateur
PT15    02/10/2003 SB V_42  Affichage des ongles si gestion paramsoc des combos libres
PT16    06/10/2003 SB V_42  Erreur SQL base ORACLE, restructuration des joints
PT17  : 27/11/2003 SB V_50  Suppression des variables globales appelés dans pgedtetat
PT18  : 04/05/2004 SB V_50  FQ 11112 Editon bulletin à zéro
PT19  : 05/05/2004 SB V_50  FQ 11202 En mode etats chaines, édition du modèle défini dans les param. soc
PT20  : 05/10/2004 PH V_50  FQ 11652 Correcttion PT18 TOUTES les lignes sont imprimées alors qu'elles ne sont pas alimentées
PT21  : 21/03/2005 SB V_60  FQ 12077 Affectation des valeurs par défaut sur le update et non sur le argument
PT22  : 26/05/2005 SB V_60  FQ 12308 Modification de l'ordre de présentation des rubriques
PT23  : 12/07/2005 SB V_60  FQ 12308 Modification de l'ordre de présentation des bulletins
PT24  : 13/02/2006 SB V_65  FQ 12451 Ajout option edition bulletin salarié par défaut
PT25  : 21/04/2006 EPI V_65  Ajout sélection salarié par le processus
PT26  : 02/05/2006 SB V_65 FQ 12579 Refonte des tri
PT27  : 31/10/2006 SB V_70  FQ 13344 Ajout controle saisie date d'affichage
PT28  : 06/03/2007 FC V_70  Gestion des éditions d'originaux, de duplicatas et de specimens
                            Ajout de la récupération du champ PPU_TOPEDITION pour le mettre à jour après l'édition
                            Tout ceci n'est valable que pour l'état PBP/PBD
PT29  : 16/04/2007 FC V_72  FQ 14089 Pouvoir visualiser tous les bulletins du salarié
PT30  : 27/04/2007 FC V_72  Appel d'une procédure de vérification de la validité du paramétrage des populations
PT31  : 03/05/2007 FC V_72  Prendre en compte le modèle de bulletin par défaut du salarié si édition depuis la saisie
PT32  : 06/06/2007 MF V_72  FQ 14329 : adaptation pour états chaînés
PT33  : 07/06/2007 MF V_72  FQ 14255 : Correction visualisation des bulletins (clique droit en saisie bulletin)
PT34  : 09/07/2007 FC V_72  FQ 14535 Historisation population
PT35  : 28/08/2007 FC V_80  appel de la dll pour les fonctions @
PT36  : 05/12/2007 PH V_80  Mise en place des codes OMR
PT37  : 07/12/2007 FC V_80  Mise en commentaire de l'utilisation de la dll en CWAS car non testé
PT38  : 27/12/2007 FC V_81  FQ 14737 Touches M S T A que les dates de paie
PT39  : 02/01/2008 FC V_81  FQ 15044 Rajout champ PHB_COTREGUL
}
unit  UTOFPGEditBulletin;

interface
uses StdCtrls,Controls,Classes,sysutils,ComCtrls,
{$IFDEF EAGLCLIENT}
     eQRS1,UTOB,
{$ELSE}
     {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}QRS1,
{$ENDIF}
     HCtrls,HEnt1,HMsgBox,UTOF,HQry,UTOFPGEtats
{$IFNDEF CPS1}
  ,PGPOPULOUTILS  //PT30
{$ENDIF}
{$IFDEF EAGLCLIENT}
     ,utileAGL
     ,uLanceProcess
     ,uHTTP
{$ELSE}
     ,uEdtComp
{$ENDIF}
     ;

Type
     TOF_PGEDIT_BUL = Class (TOF_PGEtats)  //PT10
       public
       procedure OnArgument(Arguments : String ) ; override ;
       procedure OnUpdate; override ;
       procedure OnLoad; override;
       procedure OnClose; override;    //PT35
       private
       Origine,Etab,CodeSal,DebPer,FinPer : string;
       DDSaisBul,DFSaisBul : TDateTime;   //PT3
       ModeEdition:String;//PT28
       {procedure Change(Sender: TObject);  PT26 }
       procedure ExitEdit(Sender: TObject);
       procedure AffectPeriodePaie(Sender: TObject);
       procedure ForceLibellePaie(Sender: TObject);
       procedure ChangeCriterePeriode(Sender : Tobject);
       procedure AffectDatePaie (DatePaie : TDateTime);  //PT10
       procedure AfficheIntitule; { PT21 }
       procedure OnChangeTri (Sender : Tobject); { PT26 }
       function  PgRendOrdebyEtatBul : string;   { PT26 }
     END ;

implementation

uses Vierge,PgEditOutils,PgEditOutils2,EntPaie,PgOutils,PGoutils2,
  ParamSoc //PT28
  ,BullOLE,EdtEtat,BullEdCalc //PT35
  ;

//DEB PT37
{$IFDEF DLL}
//DEB PT35
{$IFDEF EAGLCLIENT}
function CalcPaieServer(sf,sp : string) : variant ;
var T, TResponse : TOB;
begin
  T := TOB.Create ('',nil,-1);
  try
    { Ouverture du serveur }
    TResponse := AppServer.Request('Paie.PaieEdOpen','',T,'','');
    if TResponse = nil then PGIBox('Paie ( Ouverture ) : serveur d''éditions introuvable');
    if TResponse <> nil then TResponse.Free;
  finally
    T.Free;
  end;
end;
{$ENDIF}
//FIN PT35
{$ENDIF DLL}
//FIN PT37

procedure TOF_PGEDIT_BUL.OnArgument(Arguments: String);
var
    Defaut  : THEdit;
    Arg,Min,Max,ExerPerEncours{,DebPer,FinPer PT7-1 }: string;
    check :TCheckBox;
    Combo : THValComboBox;
    BullCompl : Boolean;  //PT6
    LeSalarie,Opt : String;  // PT25
    i, j : Integer; { PT26 }
    Q : TQuery; //PT30
begin
inherited ;
Arg:=Trim(Arguments);
//Edition à partir du menu,de la saisie du bulletin ou des états chaînés
//DEB PT10 Intégration du traitement d'états chaînés
DFSaisBul:=idate1900;
// PT25 début
// if Pos('MENU',Arg)>0 then Origine:='MENU'
if Pos('MENU',Arg)>0 then
begin
     Origine:='MENU';
     Opt := ReadTokenSt(Arguments);
     if Opt = 'MENU' then
     begin
       LeSalarie := ReadTokenSt(Arguments);
       ModeEdition := ReadTokenSt(Arguments);
     end
     else
       ModeEdition := ReadTokenSt(Arguments);
end
// PT25 fin
  else
    if Pos('CHAINES',Arg)>0 then
      Begin
      Origine:=ReadTokenSt(Arg);
      //DEB PT28 Si on distingue les différents modes d'édition, comme on n'a pas la main
      //         on fait des originaux systématiquement.
      //         Vérifier cependant que l'utilisateur a le droit d'éditer des originaux
      if GetParamSocSecur('SO_PGGESTORIDUPSPE',False) then
        ModeEdition := 'ORIGINAL';
      //FIN PT28
      if Trim(Arg)<>'' then
        Begin
        DDSaisBul:=StrToDate(ReadTokenSt(Arg));
        DFSaisBul:=StrToDate(ReadTokenSt(Arg));
        AffectDatePaie(DFSaisBul);
        AffectPeriodePaie(nil);
        End;
      End
    Else
    if Pos('BULL',Arg)>0 then
      Begin
      Etab:=ReadTokenSt(Arg);
      CodeSal:=ReadTokenSt(Arg);
      DDSaisBul:=StrToDate(ReadTokenSt(Arg));   //PT3
      DFSaisBul:=StrToDate(ReadTokenSt(Arg));   //PT3
      Origine:=ReadTokenSt(Arg);
      BullCompl:=(ReadTokenSt(Arg)='X'); //DEB    PT6
      ModeEdition := ReadTokenSt(Arg); //PT28
      SetControlChecked('PPU_BULCOMPL',BullCompl);   //FIN PT6
      End
    else
    //DEB PT29
    if Pos('VISUBUL',Arg)>0 then
      begin
      Etab:=ReadTokenSt(Arg);
      CodeSal:=ReadTokenSt(Arg);
      DFSaisBul:=StrToDate(ReadTokenSt(Arg));   //PT3
      Origine:=ReadTokenSt(Arg);
      BullCompl:=(ReadTokenSt(Arg)='X'); //DEB    PT6
      ModeEdition := ReadTokenSt(Arg); //PT28
      SetControlChecked('PPU_BULCOMPL',BullCompl);   //FIN PT6
      end;
    //FIN PT29

//FIN PT10
//DEB PT9
{$IFDEF CCS3}
SetControlVisible('PPU_BULCOMPL', False);
SetControlChecked('PPU_BULCOMPL', False);
{$ENDIF}
//FIN PT9

//DEB PT28
if (ModeEdition = 'ORIGINAL') then
begin
  SetControlVisible('ORIGINAL', True);
  SetControlChecked('ORIGINAL', True);
  SetControlVisible('DUPLICATA', False);
  SetControlChecked('DUPLICATA', False);
  SetControlVisible('SPECIMEN', False);
  SetControlChecked('SPECIMEN', False);
end
else if (ModeEdition = 'DUPLICATA') then
begin
  SetControlVisible('ORIGINAL', False);
  SetControlChecked('ORIGINAL', False);
  SetControlVisible('DUPLICATA', True);
  SetControlChecked('DUPLICATA', True);
  SetControlVisible('SPECIMEN', False);
  SetControlChecked('SPECIMEN', False);
end
else if (ModeEdition = 'SPECIMEN') then
begin
  SetControlVisible('ORIGINAL', False);
  SetControlChecked('ORIGINAL', False);
  SetControlVisible('DUPLICATA', False);
  SetControlChecked('DUPLICATA', False);
  SetControlVisible('SPECIMEN', True);
  SetControlChecked('SPECIMEN', True);
end
else
begin
  SetControlVisible('ORIGINAL', False);
  SetControlChecked('ORIGINAL', False);
  SetControlVisible('DUPLICATA', True);
  SetControlChecked('DUPLICATA', False);
  SetControlEnabled('DUPLICATA', True);
  SetControlVisible('SPECIMEN', False);
  SetControlChecked('SPECIMEN', False);
end;
//FIN PT28

if VH_Paie.PgBulDefaut<>'' then
   TFQRS1(Ecran).CodeEtat:=VH_Paie.PgBulDefaut;
AfficheIntitule; { PT21 }

if GetParamSocSecur ('SO_PGBULSOUSPLI', FALSE) then // PT36
 V_PGI.MiseSousPli := TRUE;

//DEB PT35
//DEB PT37
{$IFDEF DLL}
{$IFDEF EAGLCLIENT}
 ProcCalcEdt := nil;
 CalcPaieServer('','') ;
{$ELSE}
//  ProcCalcEdt := PGEdtEtat ;
  ProcCalcEdt := CalcOLEEtatBull ;
{$ENDIF}
{$ELSE}
  ProcCalcEdt := CalcOLEEtatBull ;
{$ENDIF DLL}
//FIN PT37
//FIN PT35

// d PT33
// Qd clique droit en saisie bulletin la modification des critères "dates" est possible
if (Origine <> 'VISUBUL') then
begin
  SetControlEnabled('DATEDEB',False);
  SetControlEnabled('DATEDEB_',False);   //PT38
end;
// f PT33

//PT10 Ajout traitement Etats chaînés
if (origine ='MENU') or (origine ='VISUBUL') or ((Origine='CHAINES') AND (DFSaisBul=idate1900)) then
  Begin     //Valeur par défaut
    //PT25 début
    // RecupMinMaxTablette('PG','SALARIES','PSA_SALARIE',Min,Max);
    If LeSalarie = '' then
    begin
      RecupMinMaxTablette('PG','SALARIES','PSA_SALARIE',Min,Max);
    end
    else
    begin
      Min := LeSalarie;
      Max := LeSalarie;
    end;
    //PT25 fin
    Defaut:=ThEdit(getcontrol('PPU_SALARIE'));
    If Defaut<>nil then Begin Defaut.text:=Min; Defaut.OnExit:=ExitEdit; End;
    Defaut:=ThEdit(getcontrol('PPU_SALARIE_'));
    If Defaut<>nil then Begin Defaut.text:=Max;  Defaut.OnExit:=ExitEdit; End;
    RecupMinMaxTablette('PG','ETABLISS','ET_ETABLISSEMENT',Min,Max);
    Defaut:=ThEdit(getcontrol('PPU_ETABLISSEMENT'));
    If Defaut<>nil then Defaut.text:=Min;
    Defaut:=ThEdit(getcontrol('PPU_ETABLISSEMENT_'));
    If Defaut<>nil then Defaut.text:=Max;
    Combo:=THValComboBox(GetControl('CBMOIS'));
    if combo<>nil then Combo.OnChange:=AffectPeriodePaie;
    Combo:=THValComboBox(GetControl('CBANNEE'));
    if combo<>nil then Combo.OnChange:=AffectPeriodePaie;
    if RendPeriodeEnCours (ExerPerEncours,DebPer,FinPer) then
      begin
      AffectDatePaie(StrToDate(DebPer)); //PT10
      end;
  //Rend visible les Org Stat et code Stat en fonction des ParamSoc
  VisibiliteChamp (Ecran);
  VisibiliteChampLibre (Ecran);
  { DEB PT15 }
  SetControlProperty('TBCOMPLEMENT','Tabvisible',(VH_Paie.PGNbreStatOrg>0) OR (VH_Paie.PGLibCodeStat<>''));
  SetControlProperty('TBCHAMPLIBRE','Tabvisible',(VH_Paie.PgNbCombo>0));
  { FIN PT15 }

  //Evenement On Click sur les ruptures
 {  PT26 Mise en commentaire
  CTrav1:=TCheckBox(GetControl('CN1'));  if CTrav1<>nil then CTrav1.OnClick:=Change;
  CTrav2:=TCheckBox(GetControl('CN2'));  if CTrav2<>nil then CTrav2.OnClick:=Change;
  CTrav3:=TCheckBox(GetControl('CN3'));  if CTrav3<>nil then CTrav3.OnClick:=Change;
  CTrav4:=TCheckBox(GetControl('CN4'));  if CTrav4<>nil then CTrav4.OnClick:=Change;
  CTrav5:=TCheckBox(GetControl('CN5'));  if CTrav5<>nil then CTrav5.OnClick:=Change;
  CEtab:=TCheckBox(GetControl('CETAB')); if Cetab<>nil then Cetab.OnClick:=Change;
  Alpha:=TCheckBox(GetControl('CALPHA'));if Alpha<>nil then  Alpha.OnClick:=Change;
  Check:=TCheckBox(GetControl('CL1'));   if Check<>nil then Check.OnClick:=Change;
  Check:=TCheckBox(GetControl('CL2'));   if Check<>nil then Check.OnClick:=Change;
  Check:=TCheckBox(GetControl('CL3'));   if Check<>nil then Check.OnClick:=Change;
  Check:=TCheckBox(GetControl('CL4'));   if Check<>nil then Check.OnClick:=Change;}
  Check:=TCheckBox(GetControl('CKFORCELIB'));   if Check<>nil then Check.OnClick:=ForceLibellePaie;
  Check:=TCheckBox(GetControl('CKMENS'));   if Check<>nil then Check.OnClick:=ChangeCriterePeriode;//PT7-1

{ DEB PT26 }
  for j := 1 to 4 do
  begin
    Combo := ThValComboBox(GetControl('THVALTRI' + IntToStr(j)));
    if Assigned(Combo) then
     Begin
     Combo.onchange := OnChangeTri;
     Combo.Itemindex := 0;
     Combo.Items.Add('<<Aucune>>');          Combo.Values.Add('');
     Combo.Items.Add('Etablissement');       Combo.Values.Add('PPU_ETABLISSEMENT');
     Combo.Items.Add('Salarié');             Combo.Values.Add('PPU_SALARIE');
     Combo.Items.Add('Nom du salarié');      Combo.Values.Add('PPU_LIBELLE,PPU_PRENOM');
     Combo.Items.Add('Mode de règlements');  Combo.Values.Add('PPU_PGMODEREGLE');
     For i :=1 To VH_Paie.PGNbreStatOrg Do
       Begin
       IF (i= 1) AND (VH_Paie.PGLibelleOrgStat1 <> '') then Combo.Items.Add(VH_Paie.PGLibelleOrgStat1)
       else IF (i= 2) AND (VH_Paie.PGLibelleOrgStat2 <> '') then Combo.Items.Add(VH_Paie.PGLibelleOrgStat2)
       else IF (i= 3) AND (VH_Paie.PGLibelleOrgStat3 <> '') then Combo.Items.Add(VH_Paie.PGLibelleOrgStat3)
       else IF (i= 4) AND (VH_Paie.PGLibelleOrgStat4 <> '') then Combo.Items.Add(VH_Paie.PGLibelleOrgStat4);
       Combo.Values.Add('PPU_TRAVAILN'+IntToStr(i));
       End;
     if VH_Paie.PGLibCodeStat <> '' then
       Begin
       Combo.Items.Add(VH_Paie.PGLibCodeStat);
       Combo.Values.Add('PPU_CODESTAT');
       End;
     For i :=1 To VH_Paie.PgNbCombo Do
       Begin
       IF (i= 1) AND (VH_Paie.PgLibCombo1 <> '') then Combo.Items.Add(VH_Paie.PgLibCombo1)
       else IF (i= 2) AND (VH_Paie.PgLibCombo2 <> '') then Combo.Items.Add(VH_Paie.PgLibCombo2)
       else IF (i= 3) AND (VH_Paie.PgLibCombo3 <> '') then Combo.Items.Add(VH_Paie.PgLibCombo3)
       else IF (i= 4) AND (VH_Paie.PgLibCombo4 <> '') then Combo.Items.Add(VH_Paie.PgLibCombo4);
       Combo.Values.Add('PPU_LIBREPCMB'+IntToStr(i));
       End;
     End;
  end;
{ FIN PT26 }

    //DEB PT29
    if Origine='VISUBUL' then
    Begin
      SetControlText('PPU_SALARIE',CodeSal);
      SetControlText('PPU_SALARIE_',CodeSal);
      SetControlText('PPU_ETABLISSEMENT',Etab);
      SetControlText('PPU_ETABLISSEMENT_',Etab);
      SetControlEnabled('PPU_SALARIE',False);
      SetControlEnabled('PPU_SALARIE_',False);
      SetControlEnabled('PPU_ETABLISSEMENT',False);
      SetControlEnabled('PPU_ETABLISSEMENT_',False);
//      SetControlEnabled('DATEDEB',True);
//      SetControlEnabled('DATEFIN',True);
      AffectDatePaie(DFSaisBul);
      AffectPeriodePaie(nil);
      TFQRS1(Ecran).FDuplicata := False;
    End;
    //FIN PT29

  End
  Else
    if Origine='BULL' then
      Begin
      SetControlText('PPU_SALARIE',CodeSal);
      SetControlText('PPU_SALARIE_',CodeSal);
      AffectDatePaie(DFSaisBul); //PT10 utilisation de la procédure d'affectation des dates
      AffectPeriodePaie(nil);
      //Remise en forme de la fenetre d'edition
      //DEB PT28
      if (ModeEdition = 'SPECIMEN') then
        TFQRS1(Ecran).FDuplicata := True     // Permet d'avoir le texte imprimé en travers
      else
        TFQRS1(Ecran).FDuplicata := False;
      //FIN PT28

      TFQRS1(Ecran).Dock971.Visible:=False;
      TFQRS1(Ecran).bAgrandir.Click;
      TFQRS1(Ecran).bAgrandir.Visible:=False;
      TFQRS1(Ecran).bvalider.click;
      End
  else
    //DEB PT29
    if Origine='VISUBUL' then
    Begin
      SetControlText('PPU_SALARIE',CodeSal);
      SetControlText('PPU_SALARIE_',CodeSal);
      SetControlText('PPU_ETABLISSEMENT',Etab);
      SetControlText('PPU_ETABLISSEMENT_',Etab);
      SetControlEnabled('PPU_SALARIE',False);
      SetControlEnabled('PPU_SALARIE_',False);
      SetControlEnabled('PPU_ETABLISSEMENT',False);
      SetControlEnabled('PPU_ETABLISSEMENT_',False);
      SetControlEnabled('DATEDEB',True);
      SetControlEnabled('DATEDEB_',True);    //PT38
      AffectDatePaie(DFSaisBul);
      AffectPeriodePaie(nil);
      TFQRS1(Ecran).FDuplicata := False;
    End;
    //FIN PT29

  //DEB PT30
{$IFNDEF CPS1}
  //S'il existe un parametre associé (Pour les zones libres bulletin) , appeler la fonction de test de validité des populations
  Q := Opensql ('SELECT PGO_PGPARAMETRE FROM PGPARAMETRESASSOC WHERE PGO_TYPEPARAMETRE = "POP"',true);
  if not Q.Eof then
  begin
    if not CanUsePopulation('PAI') then
      PGIBox(TraduireMemoire('La population "PAI" utilisée dans le paramétrage des zones libres bulletin n''est pas valide'));
  end;
  Ferme(Q);
{$ENDIF}
  //FIN PT30
End;

{PT26  Mise en commentaire procedure TOF_PGEDIT_BUL.Change(Sender: TObject);
begin
BloqueChampRupture(Ecran);
BloqueChampLibre(Ecran);
if TCheckBox(Sender).Name='CALPHA' then //PT7-2
  AffectCritereAlpha(Ecran,TCheckBox(Sender).Checked,'PPU_SALARIE','PPU_LIBELLE');
end;}

procedure TOF_PGEDIT_BUL.ExitEdit(Sender: TObject);
var edit : thedit;
begin
edit:=THEdit(Sender);
if edit <> nil then
    if (VH_Paie.PgTypeNumSal='NUM') and (length(Edit.text)<11) and (isnumeric(edit.text)) then
    edit.text:=AffectDefautCode(edit,10);
end;

procedure TOF_PGEDIT_BUL.OnUpdate;
var
SQL,AnneeOk,StDate : String;
cbMois,cbAnnee :THValComboBox;
DebutDateFin,FinDateFin,DFPaie : String;
Temp,Tempo,Critere : String;
Pages:TPageControl;
x : integer;
Q : TQuery;
begin
inherited;
{ DEB PT27 }
if (isvaliddate(GetControlText('DATEDEB')))  and (isvaliddate(GetControlText('DATEDEB_')))  then //PT38
   if StrToDate(GetControlText('DATEDEB')) > StrToDate(GetControlText('DATEDEB_')) then          //PT38
           Begin
           PgiInfo('Les dates d''affichage de paie sont incohérentes.', TFQRS1(Ecran).caption);
           SetControlText('DATEDEB_',GetControlText('DATEDEB'));                                 //PT38
           End;
{ FIN PT27 }

AfficheIntitule; { PT21 }

// PT32  L'affectation de TFQRS1(Ecran).CodeEtat est déjà réalisée sur le OnArgument
//       Pour les états chaînés cette affectation est remplacée par le code état
//       du filtre utilisé si CritEdtChaine.FiltreUtilise est renseigné
// PT32 if (VH_Paie.PgBulDefaut<>'') AND (Origine='CHAINES') then { PT19 }
// PT32    TFQRS1(Ecran).CodeEtat:=VH_Paie.PgBulDefaut;

//DEB PT31
if (Origine = 'BULL') then
begin
  Q := OpenSql('SELECT PSA_ETATBULLETIN FROM SALARIES WHERE PSA_SALARIE = "' + CodeSal + '"', True);
  if not Q.eof then
    TFQRS1(Ecran).CodeEtat := Q.FindField('PSA_ETATBULLETIN').AsString;
  Ferme(Q);
end;
//FIN PT31

DFPaie:=UsDateTime(idate1900);
Pages:=TPageControl(GetControl('Pages'));
Temp:=RecupWhereCritere(Pages);
tempo:=''; critere:='';
x:=Pos('(',Temp);
if x>0 then Tempo:=copy(Temp,x,(Length(temp)-5));
if tempo<>'' then critere:='AND '+Tempo;

if GetControlText('CKMENS')='-' then     //PT7-1 Ajout Cond.
  Begin
  cbMois:=THValComboBox(GetControl('CBMOIS'));
  cbAnnee:=THValComboBox(GetControl('CBANNEE'));
  if (cbMois<>nil) and (cbAnnee<>nil)  then
    begin
    ControlMoisAnneeExer(cbMois.value, RechDom ('PGANNEESOCIALE',cbAnnee.Value,FALSE),AnneeOk );
// d PT33
    if (origine = 'VISUBUL') then
      DebutDateFin := UsDateTime(StrToDate(GetControlText('DATEDEB'))) 
    else
// f PT33
      DebutDateFin:=UsDateTime(EncodeDate(StrToInt(AnneeOk),StrToInt(cbMois.value),1));
    FinDateFin:=UsDateTime(FindeMois(EncodeDate(StrToInt(AnneeOk),StrToInt(cbMois.value),1)));
    end;
  //DEB PT3  //PT10 Affect date d'edition
//PT33  if (origine='MENU') or (origine='CHAINES') then
  if (origine='MENU') or (origine='CHAINES') or (origine='VISUBUL') then
    Begin
    StDate:=' AND PPU_DATEFIN>="'+DebutDateFin+'" AND PPU_DATEFIN<="'+FinDateFin+'"';
    DFPaie:=FinDateFin;  //PT12-2
    End
  else
    Begin
    StDate:=' AND PPU_DATEDEBUT="'+USDateTime(DDSaisBul)+'" AND PPU_DATEFIN="'+UsDateTime(DFSaisBul)+'"';
    DFPaie:=UsDateTime(DFSaisBul); //PT12-2
    End;
  //FIN PT3
  End
else        //PT7-1 Ajout Cond.
  Begin
  StDate:='';
  If (IsValidDate(GetControlText('PPU_DATEFIN'))) then
     DFPaie:=UsDateTime(StrToDate(GetControlText('PPU_DATEFIN')));  //PT12-2
  End;
{ DEB PT12-2 on exclut les salariés sortis dans les sessions de paie }
if (GetControltext('CKSORTIE')='X') and (DFPaie<>UsdateTime(idate1900)) then
  Critere:=Critere+' AND (PPU_DATESORTIE IS NULL '+
                      'OR PPU_DATESORTIE="'+Usdatetime(idate1900)+'" '+
                      'OR PPU_DATESORTIE>"'+DFPaie+'")';
{ FIN PT12-2 }

{ DEB PT24 }
if (GetControltext('CKEDITSALBULDEFAUT')='X') then
    Critere:=Critere+' AND PPU_SALARIE IN (SELECT PSA_SALARIE FROM SALARIES WHERE PSA_ETATBULLETIN="'+TFQRS1(Ecran).CodeEtat+'") ';
{ FIN PT24 }


{Utilisé pour test développement états chaînés
TFQRS1(Ecran).WhereSQL:=RendRequeteSQLEtat('BULLETIN','',Critere+StDate,OrderBy);
}
SQL:='SELECT PPU_ETABLISSEMENT,PPU_SALARIE,PPU_DATEDEBUT,PPU_DATEFIN, '+
'PPU_EDTDEBUT,PPU_EDTFIN,PPU_DATEENTREE,PPU_DATESORTIE,PPU_CIVILITE,'+   //PT8 Ajout champ civilité
'PPU_LIBELLE,PPU_PRENOM,PPU_AUXILIAIRE,'+
'PPU_ADRESSE1,PPU_ADRESSE2,PPU_ADRESSE3,PPU_CODEPOSTAL,PPU_VILLE,'+
'PPU_NUMEROSS,PPU_CONVENTION,PPU_CODEEMPLOI,PPU_INDICE,PPU_NIVEAU,'+
'PPU_QUALIFICATION,PPU_COEFFICIENT,PPU_PGMODEREGLE,'+ {PT2}
'PPU_TRAVAILN1,PPU_TRAVAILN2,PPU_TRAVAILN3,PPU_TRAVAILN4,PPU_CODESTAT,'+
'PPU_LIBELLEEMPLOI,PPU_PAYELE,PPU_CHEURESTRAV,PPU_CBRUT,PPU_CNETAPAYER,'+
'PPU_CBRUTFISCAL,PPU_CNETIMPOSAB,PPU_CCOUTPATRON,PPU_CCOUTSALARIE,PPU_DATEANCIENNETE,'+
'PPU_DENOMINTRENT,PPU_NUMERATTRENT,'+ //PT14
'PPU_TOPEDITION,PPU_USEREDIT,'+//PT28 //PT34
'PHB_NATURERUB,PHB_RUBRIQUE,PHB_LIBELLE,PHB_BASEREM,PHB_TAUXREM,PHB_COEFFREM,'+
'PHB_MTREM,PHB_IMPRIMABLE,PHB_BASECOT,PHB_ORDREETAT,PHB_TAUXSALARIAL,'+
'PHB_MTSALARIAL,PHB_TAUXPATRONAL,PHB_MTPATRONAL,PHB_BASEREMIMPRIM,'+
'PHB_TAUXREMIMPRIM,PHB_COEFFREMIMPRIM,PHB_BASECOTIMPRIM,PHB_TAUXSALIMPRIM,'+
'PHB_TAUXPATIMPRIM,PHB_ORGANISME,PHB_SENSBUL,PHB_OMTSALARIAL,ET_JURIDIQUE, '+     //PT12-1 PT23
'ET_LIBELLE,ET_ADRESSE1,ET_ADRESSE2,ET_CODEPOSTAL,ET_VILLE,ET_SIRET,ET_APE, '+
'POG_LIBEDITBULL,PHB_COTREGUL '+        //PT13 Ajout Champ         //PT39
'FROM PAIEENCOURS,HISTOBULLETIN '+ //DEB PT16
'LEFT JOIN ETABLISS ON ET_ETABLISSEMENT=PHB_ETABLISSEMENT '+
'LEFT JOIN ORGANISMEPAIE ON POG_ETABLISSEMENT=PHB_ETABLISSEMENT AND PHB_ORGANISME=POG_ORGANISME '+ //PT13 Ajout Join
'WHERE PPU_ETABLISSEMENT=PHB_ETABLISSEMENT AND PPU_SALARIE=PHB_SALARIE '+
'AND PPU_DATEDEBUT=PHB_DATEDEBUT AND PPU_DATEFIN=PHB_DATEFIN '+ //FIN PT16
'AND PHB_NATURERUB<>"BAS" '+
'AND ((PHB_MTSALARIAL<>0 OR PHB_MTPATRONAL<>0 OR PHB_MTREM<>0 OR PHB_RUBRIQUE like "%.%" '+
'OR (PHB_BASEREM<>0 AND PHB_BASEREMIMPRIM="X" AND PHB_TAUXREMIMPRIM="-" AND PHB_COEFFREMIMPRIM="-")) '+//PT5 //PT11 ajout clause // PT20 Ajout condition
'OR ((PPU_CNETAPAYER=0 AND PPU_CBRUT=0) AND (PHB_BASEREM<>0 AND PHB_BASEREMIMPRIM="X" AND PHB_TAUXREMIMPRIM="-" AND PHB_COEFFREMIMPRIM="-"))) '+  { PT18 Ajout clause }
'AND PHB_IMPRIMABLE="X" AND PHB_ORDREETAT>0 '+Critere+ StDate + PgRendOrdebyEtatBul; { PT22 } { PT26 }

TFQRS1(Ecran).WhereSQL:=SQL;     { PT26 }

//DEB PT28
if (ModeEdition = 'SPECIMEN') then
  TFQRS1(Ecran).FDuplicata := True     // Permet d'avoir le texte imprimé en travers
else
  TFQRS1(Ecran).FDuplicata := False;
//FIN PT28

//PGLanceBulletin:=True; PT17
end;


procedure TOF_PGEDIT_BUL.AffectPeriodePaie(Sender: TObject);
var
Mois,Annee,AnneeOk : string;
DebPaie,FinPaie : TDateTime;
begin
Mois:=GetControlText('CBMOIS');
Annee:=GetControlText('CBANNEE');
if (Mois='') or (Annee='') then exit;
ControlMoisAnneeExer(Mois, RechDom ('PGANNEESOCIALE',Annee,FALSE),AnneeOk );
if AnneeOk<>'' then
  Begin
  DebPaie:=EncodeDate(StrToInt(AnneeOk),StrToInt(Mois),1);
  FinPaie:=FindeMois(EncodeDate(StrToInt(AnneeOk),StrToInt(Mois),1));
  SetControlText('DATEDEB',DateToStr(DebPaie));
  SetControlText('DATEDEB_',DateToStr(FinPaie)); //PT38
  End;
end;

procedure TOF_PGEDIT_BUL.ForceLibellePaie(Sender: TObject);
begin
SetControlEnabled('DATEDEB',TCheckBox(Sender).Checked);
SetControlEnabled('DATEDEB_',TCheckBox(Sender).Checked); //PT38
end;

//DEB PT7-1
procedure TOF_PGEDIT_BUL.ChangeCriterePeriode(Sender: Tobject);
begin
if TCheckBox(Sender).Checked then
  Begin
  SetControlVisible('CBMOIS',False);
  SetControlVisible('CBANNEE',False);
  SetControlProperty('DD','Name','PPU_DATEDEBUT');
  SetControlProperty('DF','Name','PPU_DATEFIN');
  SetControlVisible('PPU_DATEDEBUT',True);
  SetControlVisible('PPU_DATEFIN',True);
  SetControlVisible('TDF',True);
  SetControlText('PPU_DATEDEBUT',DebPer);
  SetControlText('PPU_DATEFIN',FinPer);
  End
else
  Begin
  SetControlVisible('CBMOIS',True);
  SetControlVisible('CBANNEE',True);
  SetControlProperty('PPU_DATEDEBUT','Name','DD');
  SetControlProperty('PPU_DATEFIN','Name','DF');
  SetControlVisible('DD',False);
  SetControlVisible('DF',False);
  SetControlVisible('TDF',False);
  SetControlText('PPU_DATEDEBUT',DateToStr(idate1900));
  SetControlText('PPU_DATEFIN',DateToStr(idate1900));
  End;
end;
//FIN PT7-1
procedure TOF_PGEDIT_BUL.OnLoad;
begin
  inherited;
//PT10 Traitement Etats chaînés
if Origine='CHAINES' then AffectDatePaie(DFSaisBul);
end;

//DEB PT10 procedure d'affectation des dates de paie
procedure TOF_PGEDIT_BUL.AffectDatePaie(DatePaie: TDateTime);
var YY,MM,JJ : WORD;
    Mois,ExerPerEncours : string;
begin
if DatePaie>idate1900 then
  Begin
  DecodeDate(DatePaie,YY,MM,JJ);
  Mois:=IntToStr(MM);
  if Length(Mois)=1 then Mois:='0'+Mois;
  SetControlText('CBMOIS',Mois);
  ExerPerEncours:=RendExerciceCorrespondant(DatePaie);
  if (ExerPerEncours='') then
    PgiBox('Aucun exercice ne correspond à cette session de paie','Exercice social erronée');
  SetControlText('CBANNEE',ExerPerEncours);
  End;
end;
//FIN PT10
{ DEB PT21 }
procedure TOF_PGEDIT_BUL.AfficheIntitule;
Var Q : TQuery;
begin
//Affichage des zones libres
SetControlChecked('CKEURO',VH_Paie.PGTenueEuro);
SetControlChecked('PGEDITORG1',VH_Paie.PgEditOrg1);
SetControlChecked('PGEDITORG2',VH_Paie.PgEditOrg2);
SetControlChecked('PGEDITORG3',VH_Paie.PgEditOrg3);
SetControlChecked('PGEDITORG4',VH_Paie.PgEditOrg4);
SetControlChecked('PGEDITCODESTAT',VH_Paie.PgEditCodeStat);
//Affichage du taux de parité
Q:=OpenSql('SELECT D_PARITEEURO FROM DEVISE WHERE D_DEVISE="'+VH_Paie.PGMonnaieTenue+'"',True);
if not Q.eof then
  SetControlText('TAUX',Q.FindField('D_PARITEEURO').asstring)
else
  SetControlText('TAUX','1');
Ferme(Q);
//Gestion des cumuls affichés en bas de bulletin
if VH_PAIE.PGCumul01<>'' then SetControlText('XX_VARIABLECUM1',rechdom('PGCUMULPAIE',VH_PAIE.PGCumul01,False));
if VH_PAIE.PGCumul01<>'' then SetControlText('XX_VARIABLECUM2',rechdom('PGCUMULPAIE',VH_PAIE.PGCumul02,False));
if VH_PAIE.PGCumul01<>'' then SetControlText('XX_VARIABLECUM3',rechdom('PGCUMULPAIE',VH_PAIE.PGCumul03,False));
if VH_PAIE.PGCumul01<>'' then SetControlText('XX_VARIABLECUM4',rechdom('PGCUMULPAIE',VH_PAIE.PGCumul04,False));
if VH_PAIE.PGCumul01<>'' then SetControlText('XX_VARIABLECUM5',rechdom('PGCUMULPAIE',VH_PAIE.PGCumul05,False));
if VH_PAIE.PGCumul01<>'' then SetControlText('XX_VARIABLECUM6',rechdom('PGCUMULPAIE',VH_PAIE.PGCumul06,False));
 //Gestion des cumuls du mois affichés en bas de bulletin
SetControlText('LBLCUMULMOIS1',VH_Paie.PgLibCumulMois1);
SetControlText('LBLCUMULMOIS2',VH_Paie.PgLibCumulMois2);
SetControlText('LBLCUMULMOIS3',VH_Paie.PgLibCumulMois3);
end;
{ FIN PT21 }
{ DEB PT26 }
procedure TOF_PGEDIT_BUL.OnChangeTri(Sender: Tobject);
var
  Combo: THValComboBox;
  Name, Ordre: string;
  IntOrdre, i: integer;
begin
  if Sender is THValComboBox then Combo := THValComboBox(Sender) else exit;
  Name := Combo.name;
  Ordre := Copy(name, Length(name), 1);
  IntOrdre := StrToInt(Ordre) + 1;
  if Combo.value = '' then
  begin
    for i := IntOrdre to 4 do
    begin
      SetControlText('THVALTRI' + IntToStr(i), '');
      SetControlEnabled('THVALTRI' + IntToStr(i), False);
    end;
  end
  else
    SetControlEnabled('THVALTRI' + IntToStr(IntOrdre), True);
end;
{ FIN PT26 }
{ DEB PT26 }
function TOF_PGEDIT_BUL.PgRendOrdebyEtatBul: string;
Var
 i : integer;
 combo : THValComboBox;
begin
for i := 1 to 4 do
 If GetControlText('THVALTRI' + IntToStr(i))<>'' then
    begin
    Combo := ThValComboBox(GetControl('THVALTRI' + IntToStr(i)));
    if Assigned(Combo) then
      if (Combo.value <> '') AND (Pos(Combo.value,Result) = 0) then
        Result := Result + Combo.value+',';
    End;
if POS('PPU_SALARIE', Result) = 0 then Result := Result + 'PPU_SALARIE,';
Result:=' ORDER BY '+Result+'PPU_DATEDEBUT,PPU_DATEFIN,PHB_ORDREETAT,PHB_OMTSALARIAL,PHB_NATURERUB,PHB_RUBRIQUE' ; { PT23 }

end;
{ FIN PT26 }

//DEB PT35
procedure TOF_PGEDIT_BUL.OnClose;
//DEB PT37
{$IFDEF DLL}
{$IFDEF EAGLCLIENT}
  var T, TResponse : TOB;
{$ENDIF}
{$ENDIF DLL}
//FIN PT37
begin
//DEB PT37
{$IFDEF DLL}
  {$IFDEF EAGLCLIENT}
  T := TOB.Create ('',nil,-1);
  try
    TResponse := AppServer.Request('Paie.PaieEdClose','',T,'','');
    if TResponse = nil then PGIBox('Paie ( Fermeture ) : serveur d''éditions introuvable');
    TResponse.Free;
    delay(100); // on laisse le temps au serveur de se fermer
  finally
    T.Free;
  end;
  {$ENDIF}
{$ENDIF DLL}
//FIN PT37
  ProcCalcEdt:=BullCalcOLEEtat ;
  V_PGI.MiseSousPli := FALSE;  // PT36
  inherited;
end;
//FIN PT35

Initialization
registerclasses([TOF_PGEDIT_BUL]);
end.
