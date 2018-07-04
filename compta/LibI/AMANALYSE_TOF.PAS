{***********UNITE*************************************************
Auteur  ...... :
Créé le ...... : 13/05/2003
Modifié le ... : 25/04/2006
Description .. : Source TOF de la FICHE : AMANALYSE ().
Suite ........ : TEST Teamlink
Suite ........ :
Suite ........ : XVI 18/04/2006 FD 3978 Ajoute de l'analitques aux
Suite ........ : Immos S1
Suite ........ : XVI 25/04/2006 FD 3978 Ajoute de l'analitques mono
Suite ........ : ventilation pour S1
Suite ........ : MVG 12/07/2006 report des modifications de XVI
Suite ........ : BTY 09/06 FQ 16149 Appliquer les restrictions utilisateur au critère
Suite ........ : de sélection Etablissements et à la restauration de filtre
Suite ........ : MBO 05/10/2007 - FQ 21530 
Mots clefs ... : TOF;AMANALYSE
*****************************************************************}
Unit AMANALYSE_TOF ;

Interface

Uses StdCtrls,
     Controls,
     Classes,
{$IFDEF EAGLCLIENT}
     MaineAgl,
     eMul,  // BTY FQ 16149
{$ELSE}
     mul,   // BTY FQ 16149
     db,
     {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF}
     fe_main,
{$ENDIF}
{$IFDEF VER150}
   Variants,
{$ENDIF}
     forms,
     sysutils,
     ComCtrls,
     HCtrls,
     HEnt1,
     HMsgBox,
     UTOF,
     uTOB,
     Stat,
     HQry
     {$IFDEF SERIE1}
     {$ELSE}
     , Ent1
     {$ENDIF}
     ;

Type
  TOF_AMANALYSE = Class (TOF)
    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnDisplay                ; override ;
    procedure OnClose                  ; override ;
    procedure OnCancel                 ; override ;
    private
      FTob : TOB;
      function MaxTablesLibresImmo : integer;
      procedure OnChangeFFiltres(Sender : TObject); // BTY 09/06 FQ 16149
      procedure ApresChangeFFiltres;                // BTY 09/06 FQ 16149
  end ;

procedure AMLanceFiche_StatPrevisionnel;

Implementation

uses
   {$IFDEF MODENT1}
   CPProcMetier,
   {$ENDIF MODENT1}
  ImEdCalc, ImEnt
  {$IFDEF SERIE1}
  ,S1Util
  {$ELSE}
  {$ENDIF}
  ;

procedure AMLanceFiche_StatPrevisionnel;
begin
  AGLLanceFiche('AM','AMANALYSE','','','');
end;

procedure TOF_AMANALYSE.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_AMANALYSE.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_AMANALYSE.OnUpdate ;
begin
  Inherited ;
end ;

procedure TOF_AMANALYSE.OnLoad ;

var Q : TQuery;
    i : integer;
    j : integer;
    T : TOB;
    MaxTL : integer;
    stSQL : string;
    Coeff : double;
    bAnalytique : boolean;

const _NBPREVI = 20;

begin
  bAnalytique := (GetCheckBoxState('ANALYTIQUE')=cbChecked);
  FTob.ClearDetail;
  { Liste des champs à récupérer }
  stSQL := 'SELECT I_IMMO,I_LIBELLE,I_COMPTEIMMO,I_VALEURACHAT,I_DATEPIECEA,I_METHODEECO,I_METHODEFISC,I_TAUXECO,I_TAUXFISC,';
  //XVI 18/04/2006 FD 3978 début
  stSQL := stSQL + 'I_DUREEECO,I_DUREEFISC,I_LIEUGEO'
  {$IFDEF SERIE1}
  {$ELSE}
                 +',I_ETABLISSEMENT'
  {$ENDIF !SERIE1}
     ;

  if bAnalytique then
  {$IFDEF SERIE1}
  //XVI 25/04/2006 FD 3978 ajoute de la mono ventilation début
  if VS1Cpta.Y_analytiqueAdvanced then
  begin
     stSQL := stSQL + ' ,V_SECTION, V_TAUXMONTANT, S1.S_LIBELLE' ;
     if VS1Cpta.Y_NBSections>1 then
        stSQL := stSQL + ', V_SECTION2, S2.S_LIBELLE' ;
     if VS1Cpta.Y_NBSections>2 then
        stSQL := stSQL + ', V_SECTION3, S3.S_LIBELLE' ;
   end
   else
     StSQL:=stSQL+' ,I_TABLE9, 100 as V_TAUXMONTANT, S1.S_LIBELLE' ; ;
   //XVI 25/04/2006 FD 3978 ajoute de la mono ventilation fin
  {$ELSE}
     stSQL := stSQL + ' ,V_SECTION, V_TAUXMONTANT,S_SECTION,S_LIBELLE,S_AXE,X_AXE' ;
  {$ENDIF SERIE1}
  //XVI 18/04/2006 FD 3978 fin
  { Récupération des infos tables libres }
  MaxTL := MaxTablesLibresImmo;
  for i:=0 to MaxTL do stSQL := stSQL + ',I_TABLE'+IntToStr(i);
  stSQL := stSQL + ' FROM IMMO ';
  { Cas de l'analyse analytique }
  //XVI 18/04/2006 FD 3978 début
  if bAnalytique then
  begin
     //XVI 25/04/2006 FD 3978 ajoute de la mono ventilation début
    {$IFDEF SERIE1}
    if VS1Cpta.Y_analytiqueAdvanced then
    {$ELSE}
    {$ENDIF !SERIE1}
    stSQL := stSQL + ' LEFT JOIN VENTIL ON IMMO.I_IMMO=VENTIL.V_COMPTE ';
    {$IFDEF SERIE1}
    if VS1Cpta.Y_analytiqueAdvanced then
    begin
       stSQL := stSQL + ' left join SECTION S1 on VENTIL.V_SECTION=S1.S_SECTION ';
       if VS1Cpta.Y_NBSections>1 then
          stSQL := stSQL + ' left join SECTION S2 on VENTIL.V_SECTION2=S2.S_SECTION ';
       if VS1Cpta.Y_NBSections>2 then
          stSQL := stSQL + ' left join SECTION S3 on VENTIL.V_SECTION3=S3.S_SECTION ';
    end
    else
       stSQL := stSQL + ' left join SECTION S1 on IMMO.I_TABLE9=S1.S_SECTION ';
    //XVI 25/04/2006 FD 3978 ajoute de la mono ventilation fin
    {$ELSE}
    stSQL := stSQL + ' LEFT JOIN SECTION ON VENTIL.V_SECTION=SECTION.S_SECTION ';
    stSQL := stSQL + ' LEFT JOIN AXE ON SECTION.S_AXE=AXE.X_AXE ';
    {$ENDIF SERIE1}
  end;
  //XVI 18/04/2006 FD 3978 fin
  { Critères de l'utilisateur }
  stSQL := stSQL + RecupWhereCritere(TPageControl(GetControl('PAGES')));
  //XVI 18/04//2006 FD 3978 début
  {$IFDEF SERIE1}
  if BAnalytique and VS1Cpta.Y_analytiqueAdvanced then //XVI 25/04/2006 FD 3978 ajoute de la mono ventilation 
  begin
     stSQL := stSQL + ' and (S1.S_RANGSECTION="001" or S1.S_RANGSECTION is null)' ;
     if VS1Cpta.Y_NBSections>1 then
        stSQL := stSQL + ' and (S2.S_RANGSECTION="002" or S2.S_RANGSECTION is null)' ;
     if VS1Cpta.Y_NBSections>2 then
        stSQL := stSQL + ' and (S3.S_RANGSECTION="003" or S3.S_RANGSECTION is null)' ;
  end ;
  {$ELSE}
  {$ENDIF !SERIE1}
  //XVI 18/04/2006 FD 3978 fin
  { Chargement des informations de la base }
  Q := OpenSQL (stSQL ,True);
  FTob.LoadDetailDB('','','',Q,False,True);
  for i:=0 to FTob.Detail.Count - 1 do
  begin
    T := FTob.Detail[i];
    UpdatePlanInfo(VHImmo^.PlanInfo,T.GetValue('I_IMMO'),VHImmo^.Encours.Fin);
    if (bAnalytique and (not VarIsNull(T.GetValue('V_TAUXMONTANT')))) then Coeff := T.GetValue('V_TAUXMONTANT')/100
    else Coeff := 1;
    T.PutValue('I_VALEURACHAT',T.GetValue('I_VALEURACHAT')*Coeff );
    T.AddChampSupValeur('DOTATIONECO',VHImmo^.PlanInfo.DotationEco*Coeff);
    T.AddChampSupValeur('DOTATIONFISC',VHImmo^.PlanInfo.DotationFisc*Coeff);
    T.AddChampSupValeur('CUMULANTECO',VHImmo^.PlanInfo.CumulAntEco*Coeff);
    T.AddChampSupValeur('CUMULANTFISC',VHImmo^.PlanInfo.CumulAntFisc*Coeff);
    T.AddChampSupValeur('VNC',VHImmo^.PlanInfo.VNCEco*Coeff);
    for j := 1 to  _NBPREVI do
    begin
      T.AddChampSupValeur('PREVIECO_N+'+IntToStr(j),VHImmo^.PlanInfo.GetPreviEco(j)*Coeff);
      T.AddChampSupValeur('PREVIFISC_N+'+IntToStr(j),VHImmo^.PlanInfo.GetPreviFisc(j)*Coeff);
    end;
  end;
  Ferme ( Q );
  TFStat(ECRAN).LaTob := FTob;
  Inherited ;
end ;

procedure TOF_AMANALYSE.OnArgument (S : String ) ;
begin
  Inherited ;
  FTob := TOB.Create ('le_plan', nil, - 1);
  TFStat(Ecran).CritereVisible := True;

  // BTY 09/06 FQ 16149
  THValComboBox(GetControl('FFILTRES')).OnChange := OnChangeFFiltres;

  {$IFDEF SERIE1}
  SetControlVisible('ANALYTIQUE',VS1Cpta.Y_Saisie) ; //XVI 25/04/2006 auusi pour la mono-ventilation 18/04/2006 FD 3978
  {$ELSE}
  // BTY 09/06 FQ 16149
  ImGereEtablissement (THMultiValComboBox(GetControl('I_ETABLISSEMENT')) );
  {$ENDIF SERIE1}
end ;

procedure TOF_AMANALYSE.OnClose ;
begin
  FTob.Free;
  Inherited ;
end ;

procedure TOF_AMANALYSE.OnDisplay () ;
begin
  Inherited ;
end ;

procedure TOF_AMANALYSE.OnCancel () ;
begin
  Inherited ;
end ;

function TOF_AMANALYSE.MaxTablesLibresImmo: integer;
var LesLibelles : HTStringList;
    i, MaxTL : integer;
    St, stLibelle : string;
begin
  MaxTL := -1;
  LesLibelles:=HTStringList.Create ;
  {$IFDEF SERIE1}
  {$ELSE}
  GetLibelleTableLibre('I',LesLibelles);
  {$ENDIF}
  for i:=0 to LesLibelles.Count-1 do
  begin
    St:=LesLibelles.Strings[i] ;
    stLibelle := ReadTokenSt(St);
    if St = 'X' then MaxTL := i;
  end;
  LesLibelles.Free;
  Result := MaxTL;
end;

// BTY 09/06 FQ 16149
{---------------------------------------------------------------------------------------}
procedure TOF_AMANALYSE.OnChangeFFiltres(Sender: TObject);
begin
  // Evènement de l'ancêtre MUL
  // MBO - FQ 21530 - TFMUL(Ecran).FFiltresChange(Sender);
  TFSTAT(Ecran).FFiltresChange(Sender);   // l'ancêtre n'est pas un mul mais un sta

  // Revoir les établissements
  ApresChangeFFiltres;
end;

procedure TOF_AMANALYSE.ApresChangeFFiltres;
begin
{$IFNDEF SERIE1}
  ImControlEtab (THMultiValComboBox(GetControl('I_ETABLISSEMENT')) );
{$ENDIF}
end;
{---------------------------------------------------------------------------------------}



Initialization
  registerclasses ( [ TOF_AMANALYSE ] ) ;
end.
