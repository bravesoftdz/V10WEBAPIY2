{***********UNITE*************************************************
Auteur  ...... : 
Créé le ...... : 02/03/2004
Modifié le ... : 26/01/2005
Description .. : GCO - 02/03/2004
Suite ........ : -> Uniformisation de l'appel à FicheJournal en 2/3 et CWAS
Suite ........ : 
Suite ........ : VLA 26/01/2005
Suite ........ : Passage de la modif en série en CWAS
Mots clefs ... : 
*****************************************************************}
unit AGlStdcompta;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  hmsgbox, Hctrls,
  HStatus,   // MoveCur, InitMove, FinitMove
  uTOB,
{$IFDEF EAGLCLIENT}
  eFiche,   // TFFiche
  eMul,
  UtileAGL,
{$ELSE}
  {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF}
  DBGrids,
  HDB,
  Fiche,     // TFFiche
  Mul,       // TFMul
{$ENDIF}
  CPSection_TOM,  // FicheSection
  CPTiers_TOM,    // FicheTiers
  CPGeneraux_TOM, // FicheGene
  CPJournal_TOM,  // FicheJournal
  HEnt1,          // taCreatEnSerie, taModif, V_PGI
  M3FP,           // ResgisterAglProc
  AGLInit,        // TheMULQ
{$IFNDEF PGIIMMO}
{$IFNDEF CCADM}
  MZSUtil,        // ModifieEnSerie
{$ENDIF CCADM}
{$ENDIF}
  Ent1,           // MaxAxe, VH, EstComptaSansAna
  ParamSoc,       // GetParamSoc
{$IFNDEF CCADM}
  Letbatch,       // VentileGenerale, MajSoldeSection
{$ENDIF CCADM}
  UtilSais,       // MajSoldeSectionTOB
  Hqry            // THQuery
  ;

Function  AnaEstVentilable(Quelaxe : Byte; Cpte : string) : Boolean ;
{ deplacement dans ulibanalytique
procedure MajGVentil(OKgene : Boolean; LeCpte : string); // VLA 26/01/2005
}
procedure AglCreationCompte (parms: array of variant; nb: integer);
procedure AglModifCompte (parms: array of variant; nb: integer);
procedure AglEditCompte (parms: array of variant; nb: integer);
procedure AGLValiderNum(parms: array of variant; nb: integer );

implementation

uses
  {$IFDEF MODENT1}
  CPTypeCons,
  CPProcMetier,
  {$ENDIF MODENT1}
  Ulibanalytique, // REcherchePremDerAxeVentil
  SaisUtil,
{$IFNDEF PGIIMMO}
{$IFNDEF CCADM}
  PLANGENE_TOF,  // PlanGeneral
  PLANAUXI_TOF,  // PlanAuxi
  PLANSECT_TOF,  // PlanSection
  PLANJOUR_TOF,  // PlanJournal
{$ENDIF CCADM}
{$ENDIF}
(*
{$IFNDEF PGIIMMO}
  QRPlanGe,  // PlanGeneral
  QRTIER,    // PlanAuxi
  QRSECTIO,  // PlanSection
  QRPLANJO,  // PlanJournal
{$ENDIF}
*)
{$IFDEF EAGLCLIENT}

{$ELSE}
  Spin,
  Region,
{$ENDIF}
  ed_tools;


// ajout me 30/04/01
////////////////////////////////////////////////////////////////////////////////
Function AnaEstVentilable(Quelaxe : Byte; Cpte : string) : Boolean ;
var
Q1       : TQuery;
BEGIN
Result:=False ;
Q1 := Opensql ('SELECT G_VENTILABLE1,G_VENTILABLE2,G_VENTILABLE3,G_VENTILABLE4,G_VENTILABLE5 From GENERAUX Where G_GENERAL="'+Cpte+'"', TRUE);
if not Q1.EOF then
begin
     Case QuelAxe of
       0: if (Q1.FindField('G_VENTILABLE1').asstring='X') or
              (Q1.FindField('G_VENTILABLE2').asstring='X') or
              (Q1.FindField('G_VENTILABLE3').asstring='X') or
              (Q1.FindField('G_VENTILABLE4').asstring='X') or
              (Q1.FindField('G_VENTILABLE5').asstring='X') then Result:=True;
       1: if Q1.FindField('G_VENTILABLE1').asstring='X' then Result:=True ;
       2: if Q1.FindField('G_VENTILABLE2').asstring='X' then Result:=True ;
       3: if Q1.FindField('G_VENTILABLE3').asstring='X' then Result:=True ;
       4: if Q1.FindField('G_VENTILABLE4').asstring='X' then Result:=True ;
       5: if Q1.FindField('G_VENTILABLE5').asstring='X' then Result:=True ;
    End ;
end;
ferme (Q1);
END ;


////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : YMO
Créé le ...... : 01/01/1900
Modifié le ... : 06/07/2006
Description .. : Changement du mode en tacreat
Suite ........ : FQ18337 
Mots clefs ... : 
*****************************************************************}
procedure AglCreationCompte (parms: array of variant; nb: integer);
var lMode : TActionFiche ;
begin

// SBO 05/04/2007 : mise en standard de la création en série sur les comptes
{  if GetParamSocSecur('SO_CPCREATENSERIE', False)
    then lMode := taCreatEnSerie
    else lMode := taCreat ;
}
  {$IFDEF EAGLCLIENT}
  lMode := taCreat ;
  {$ELSE}
  lMode := taCreatEnSerie ;
  {$ENDIF EAGLCLIENT}

  if parms[0] = 'GENERAUX' then
    FicheGene(Nil,'','',lMode,0);

  if parms[0] = 'TIERS' then
    FicheTiers(Nil,'','',lMode,0);

  if parms[0] = 'SECTION' then
    FicheSection(Nil,'A1','',lMode,0);

  if parms[0] = 'JOURNAL' then
    FicheJournal(Nil, '','', lMode, 0);
    
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... :
Créé le ...... : 26/01/2005
Modifié le ... :   /  /
Description .. : VLA 26/01/2005
Suite ........ : Passage de ModifieEnSerie en CWAS
Mots clefs ... :
*****************************************************************}
procedure AglModifCompte (parms: array of variant; nb: integer);
var
  F : TForm;
{$IFDEF EAGLCLIENT}
  Liste   : THGrid;
{$ELSE}
  Liste   : THDBGrid;
{$ENDIF}
  Sav     : string;
  i       : integer;
  Query   : THQuery;
  InfData : string;
begin
  F:=TForm(Longint(Parms[0])) ;
  if F=Nil then exit ;
{$IFDEF EAGLCLIENT}
  Liste:=THGrid(F.FindComponent('FListe') );
{$ELSE}
  Liste:=THDBGrid(F.FindComponent('FListe') );
{$ENDIF}
  if Liste=Nil then exit;
  Query:=THQuery(F.FindComponent('Q')) ;
  if (Query=Nil) then exit;
{$IFNDEF EAGLCLIENT}
    // CA - 23/06/2003 - Pour navigation sur les enregistrements de la query
    TheMulQ:=TFMul(F).Q;
{$ENDIF}

  if (parms[2] = 'Modif') then
  begin
    if parms[3] ='GENERAUX' then
    begin
      {$IFDEF EAGLCLIENT}
        FicheGene(nil, '', parms[1], StringToAction(parms[4])(*taModif*),0) ;
      {$ELSE}
        FicheGene(Query,'',parms[1], StringToAction(parms[4]) (*taModif*) ,0) ;
      {$ENDIF}
      InfData := Query.FindField ('G_GENERAL').asstring;
    end;

    if parms[3] ='TIERS' then
    begin
      {$IFDEF EAGLCLIENT}
        FicheTiers(nil,'',parms[1], StringToAction(parms[4]) (*taModif*) ,0) ;
      {$ELSE}
        FicheTiers(Query,'',parms[1], StringToAction(parms[4]) (*taModif*) ,0) ;
      {$ENDIF}
      InfData := Query.FindField ('T_AUXILIAIRE').asstring;
    end;

    if parms[3] ='JOURNAL' then
    begin
      {$IFDEF EAGLCLIENT}
        FicheJournal(nil,'',parms[1], StringToAction(parms[4]) (*taModif*) ,0) ;
      {$ELSE}
        FicheJournal(Query,'',parms[1], StringToAction(parms[4]) (*taModif*) ,0) ;
      {$ENDIF}
      InfData := Query.FindField ('J_JOURNAL').asstring;
    end;

    if parms[3] = 'SECTION' then
    begin
      {$IFDEF EAGLCLIENT}
        FicheSection(nil,parms[4],parms[1], StringToAction(parms[5]) (*taModif*) ,0) ;
      {$ELSE}
        FicheSection(Query,parms[4],parms[1], StringToAction(parms[5]) (*taModif*) ,0) ;
      {$ENDIF}
      InfData := Query.FindField ('S_SECTION').asstring;
   end;
{$IFDEF COMPTA}
   // ajout me pour les synchronisations
   MAJHistoparam (parms[3], InfData);
{$ENDIF}
  end;

  //  Modification en série
{$IFNDEF PGIIMMO}
{$IFNDEF CCADM}
  if (parms[2] = 'Serie') then
  begin
    if (Liste.NbSelected>0) or (Liste.AllSelected) then
    BEGIN
      // parms[3] ='GENERAUX' , 'TIERS' , 'SECTIONS
      if (parms[3] = 'SECTION') then
        ModifieEnSerie(parms[3],parms[4],Liste,Query)

      //SG6 13/01/05 FQ 15224 pas d'analytiques pour les tiers
      else if (parms[3] = 'TIERS') then
      begin
        ModifieEnSerie(parms[3],'',Liste,Query);
      end
      else begin
        Sav := V_PGI.ExtendedFieldSelection;
        if (ctxPCL in V_PGI.PGIContexte) then
          if (Sav = '1') and (GetParamSocSecur('SO_CPPCLSANSANA',false)=True) then begin {Lek 100206}
            PGIInfo('L''option analytique n''est pas activée.','') ;
            exit;
          end;

          //SG6 14/01/05 FQ 15284
          if Sav = '1' then begin
            // voir pour les autres axes 3 à 5 si nécessaire
{$IFDEF CCS3}
            for i:=1 to 1 do if VH^.Cpta[AxeToFb('A'+IntToStr(i))].Attente=''
{$ELSE}
            for i:=1 to 2 do if VH^.Cpta[AxeToFb('A'+IntToStr(i))].Attente=''
{$ENDIF}
            then begin
              PGIInfo('Attention Axe '+IntToStr(i)+ ' ne possède aucun code section et n''a pas de section d''attente.' ,'') ;
              if i=2 then exit;
            end;
          end;

          {if not EstComptaSansAna then}
            if not ModifieEnSerie(parms[3],'',Liste,Query) then begin
              V_PGI.ExtendedFieldSelection := Sav;
              exit;
            end;

          if Sav = '1' then  // pour les ventilations ana
          begin
            V_PGI.ExtendedFieldSelection := Sav;
            MajGVentil (TRUE, '');
            if Liste.AllSelected then
            BEGIN
              MoveCur(False);
              Query.First;
              while Not Query.EOF do
              BEGIN
                MajGVentil (FALSE, Query.FindField ('G_GENERAL').asstring);
                Query.next;
              END;
              Liste.AllSelected:=False;
              END
            ELSE
            BEGIN
              InitMove(Liste.NbSelected,'');
              for i:=0 to Liste.NbSelected-1 do
              BEGIN
                MoveCur(False);
                Liste.GotoLeBookmark(i);
                MajGVentil (FALSE, Query.FindField ('G_GENERAL').asstring);
              END;
              Liste.ClearSelected;
              FiniMove ;
            END;
          end;
        end;
      END ;
  end;
{$ENDIF CCADM}
{$ENDIF}
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... :
Créé le ...... : 26/01/2005
Modifié le ... :   /  /
Description .. : VLA 26/01/2005
Suite ........ : Appel des éditions CWAS
Mots clefs ... :
*****************************************************************}
procedure AglEditCompte (parms: array of variant; nb: integer);
begin
{$IFNDEF PGIIMMO}
{$IFNDEF CCADM}
//{$IFDEF EAGLCLIENT}
  if parms[1] = 'GENERAUX' then CPLanceFiche_PLANGENE('')
  else if parms[1] = 'TIERS' then CPLanceFiche_PLANAUXI('')
  else if parms[1] = 'SECTION' then CPLanceFiche_PLANSECT('', '')
  else if parms[1] = 'JOURNAL' then CPLanceFiche_PLANJOUR('');
(*{$ELSE}
  if parms[1] = 'GENERAUX' then PlanGeneral('', false)
  else if parms[1] = 'TIERS' then PlanAuxi('',False)
  else if parms[1] = 'SECTION' then PlanSection('','',FALSE)
  else if parms[1] = 'JOURNAL' then PlanJournal('',FALSE);
{$ENDIF}
*)
{$ENDIF CCADM}
{$ENDIF}
end;

////////////////////////////////////////////////////////////////////////////////
procedure AGLValiderNum(parms: array of variant; nb: integer );
var
  F        : TForm;
begin
  F:=TForm(Longint(Parms[0])) ;
  TFFiche(F).ModalResult := 1;
end;

////////////////////////////////////////////////////////////////////////////////

Initialization
  RegisterAglProc( 'CreationCompte', FALSE , 1, AglCreationCompte);
  RegisterAglProc( 'ModifCompte', TRUE , 1, AglModifCompte);
  RegisterAglProc( 'EditCompte', FALSE , 1, AglEditCompte);
  RegisterAglProc( 'ValiderNum', TRUE , 1, AGLValiderNum);

finalization

end.
