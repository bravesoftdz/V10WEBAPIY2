{***********UNITE*************************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 27/06/2007
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : TBMULREVISION ()
Mots clefs ... : TOF;TBMULREVISION
*****************************************************************}
Unit UTOF_TBMULREVISION ;

Interface

Uses StdCtrls, 
     Controls, 
     Classes, 
{$IFDEF EAGLCLIENT}
     eMul,
     uTob,
     MaineAGL,
{$ELSE}
     db,
     {$IFNDEF DBXPRESS}dbtables,{$ELSE}uDbxDataSet,{$ENDIF}
     mul,
     Fe_Main,
{$ENDIF}
     forms,
     sysutils, 
     ComCtrls,
     HCtrls, 
     HEnt1,
     HTb97, 
     HMsgBox, 
     UTOF ; 

Type
  TOF_TBMULREVISION = Class (TOF)
    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnDisplay                ; override ;
    procedure OnClose                  ; override ;
    procedure OnCancel                 ; override ;

    procedure InitialiserComboBoxCompta;
    procedure BCHERCHE_OnClick(Sender: TObject);
    procedure Aff_TBFicheDossier (Sender : TObject);

  end ;

////////////////////////////////////////////////////////////////////////////////
procedure Aff_TBMulRevision;
////////////////////////////////////////////////////////////////////////////////

Implementation

uses GalOutil; //InitialiserComboGroupeConf

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 27/06/2007
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure Aff_TBMulRevision;
begin
  AGLLanceFiche('DP', 'TBMULREVISION', '', '', '');
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 27/06/2007
Modifié le ... :   /  /
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure TOF_TBMULREVISION.OnArgument (S : String ) ;
begin
  Inherited ;
  TFMul (Ecran).Caption := 'Tableau de bord révision';
  UpdateCaption (Ecran);

  THEdit(GetControl('CIR_NODOSSIER')).OnElipsisClick := Aff_TBFicheDossier;
  TToolBarButton97(GetControl('BCHERCHE')).OnClick := BCHERCHE_OnClick;

  InitialiserComboGroupeConf (THMultiValCombobox(GetControl('GROUPECONF')));
  InitialiserComboBoxCompta;
end ;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 27/06/2007
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure TOF_TBMULREVISION.OnLoad ;
begin
  Inherited ;
end ;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 27/06/2007
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure TOF_TBMULREVISION.OnDisplay () ;
begin
  Inherited ;
end ;

////////////////////////////////////////////////////////////////////////////////

procedure TOF_TBMULREVISION.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_TBMULREVISION.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_TBMULREVISION.OnUpdate ;
begin
  Inherited ;
end ;


procedure TOF_TBMULREVISION.OnClose ;
begin
  Inherited ;
end ;

procedure TOF_TBMULREVISION.OnCancel () ;
begin
  Inherited ;
end ;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 27/06/2007
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure TOF_TBMULREVISION.InitialiserComboBoxCompta;
var ChSql   : String;                               
    RSql   : Tquery;
    Indice : Integer;
    Nombre : Integer;
begin
  ChSql:='Select DTC_MILLESIME from DPTABCOMPTA where group by DTC_MILLESIME';
  try
    RSql:=Opensql (ChSql, TRUE);
    RSql.First;
    Nombre:=RSql.recordCount;

    for Indice:=0 to Nombre-1 do
    begin
      THValComboBox(GetControl('DTC_MILLESIME')).Items.AddObject (RSql.FindField ('DTC_MILLESIME').AsString,nil);
      THValComboBox(GetControl('DTC_MILLESIME')).Values.Add (RSql.FindField ('DTC_MILLESIME').AsString);
      RSql.Next;
    end;

   THValComboBox(GetControl('DTC_MILLESIME')).Value := 'N';

    Ferme (RSql);
  except
   begin
    PGIINFO ('Problème dans la table DPTABCOMPTA','Erreur');
   end;
  end;
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 27/06/2007
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure TOF_TBMULREVISION.BCHERCHE_OnClick(Sender: TObject);
var ChXXWhere : String;
begin
 ChXXWhere := GererCritereGroupeConf (THMultiValCombobox(GetControl('GROUPECONF')),TcheckBox (GetControl ('SANSGRPCONF')).Checked);
 ChXXWhere := ChXXWhere + GererCritereDivers();
 SetControlText ('XX_WHERE', chXXWHere);
 // Traitement générique
 TFMul(Ecran).BChercheClick(Sender);
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 27/06/2007
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure TOF_TBMULREVISION.Aff_TBFicheDossier(Sender: TObject);
var Chaine  : String;
    SauverDossier : String;
begin
  SauverDossier:=GetControlText('CIR_NODOSSIER');
  Chaine:=AGLLanceFiche('YY','YYDOSSIER_SEL', '','',GetControlText('CIR_NODOSSIER'));
  if Chaine <> '' then
    SetControlText ('CIR_NODOSSIER', READTOKENST(Chaine))
  else
    SetControlText('CIR_NODOSSIER', SauverDossier);
end;

////////////////////////////////////////////////////////////////////////////////

Initialization
  registerclasses ( [ TOF_TBMULREVISION ] ) ;
end.
