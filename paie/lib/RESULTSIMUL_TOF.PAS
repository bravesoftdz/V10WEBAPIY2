{***********UNITE*************************************************
Auteur  ...... : 
Créé le ...... : 30/03/2006
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : RESULTSIMUL ()
                 Liste des salariés calcul&s dans cette simulation
Mots clefs ... : TOF;RESULTSIMUL
*****************************************************************}
Unit RESULTSIMUL_TOF ;

Interface

Uses StdCtrls, 
     Controls, 
     Classes, 
{$IFNDEF EAGLCLIENT}
     db, 
     dbtables, 
     HDB,
     FE_Main,
{$else}
     uTob,
     MaineAgl,
{$ENDIF}
     forms, 
     sysutils, 
     ComCtrls,
     HCtrls,
     P5def,
     UTOF ;

Type
  TOF_RESULTSIMUL = Class (TOF)
    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnDisplay                ; override ;
    procedure OnClose                  ; override ;
    procedure OnCancel                 ; override ;
    private
    codesimul, libellesimul : string;
    procedure GrilleDblClick(Sender: TObject);

  end ;

Implementation

procedure TOF_RESULTSIMUL.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_RESULTSIMUL.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_RESULTSIMUL.OnUpdate ;
begin
  Inherited ;
end ;

{***********A.G.L.***********************************************
Auteur  ...... : NA
Créé le ...... : 06/06/2006
Modifié le ... :   /  /    
Description .. : Chargement des enregistrements
Mots clefs ... : 
*****************************************************************}
procedure TOF_RESULTSIMUL.OnLoad ;
var
stplus, stpluspop: string;
begin
  Inherited ;
  // XX_WHERE
  SetControlText('XX_WHERE','PSR_CODESIMUL ="'+codesimul+'" ');

  // Propriété PLUS pour les établissements afin d'afficher uniquemenr les établ calculés dans la simul
  Stplus := 'AND ET_ETABLISSEMENT in (select PSR_ETABLISSEMENT from resultsimul where PSR_CODESIMUL = "'+codesimul+'")';
  SetControlProperty('PSR_ETABLISSEMENT','Plus',Stplus);

  // Propriété PLUS pour les populations IDR afin d'afficher uniquement les populations calculés dans la simul
  Stpluspop := 'AND PPC_POPULATION in (select PSR_POPULATION from resultsimul where PSR_CODESIMUL = "'+codesimul+'")';
  SetControlProperty('PSR_POPULATION','Plus',Stpluspop);
end ;

{***********A.G.L.***********************************************
Auteur  ...... : NA
Créé le ...... : 06/06/2006
Modifié le ... :   /  /    
Description .. : On argument
Mots clefs ... : 
*****************************************************************}
procedure TOF_RESULTSIMUL.OnArgument (S : String ) ;
var
Letitre : String;
num : integer;
{$IFNDEF EAGLCLIENT}
    Liste: THDBGrid;
{$ELSE}
    Liste: THGrid;
{$ENDIF}
begin
  Inherited ;
  {$IFNDEF EAGLCLIENT}
   Liste := THDBGrid(GetControl('FListe'));
{$ELSE}
   Liste := THGrid(GetControl('FListe'));
{$ENDIF}
   if Liste <> nil then
   Liste.OnDblClick := GrilleDblClick;

  Codesimul := ReadTokenSt(S);
  libellesimul := ReadTokenSt(S);

  Letitre := 'Simulation :' + ' ' + Codesimul + ' ' + libellesimul;
  Ecran.caption := Ecran.caption + Letitre;

  // Recherche des libellés des zones TRAVAIL, CODESTAT
  for Num := 1 to 4 do
  begin
  VisibiliteChampSalarie(IntToStr(Num), GetControl('PSR_TRAVAILN' + IntToStr(Num)), GetControl('TPSR_TRAVAILN' + IntToStr(Num)));
  end;
  VisibiliteStat(GetControl('PSR_CODESTAT'), GetControl('TPSR_CODESTAT'));
end ;

procedure TOF_RESULTSIMUL.OnClose ;
begin
  Inherited ;
end ;

procedure TOF_RESULTSIMUL.OnDisplay () ;
begin
  Inherited ;
end ;

procedure TOF_RESULTSIMUL.OnCancel () ;
begin
  Inherited ;
end ;

{***********A.G.L.***********************************************
Auteur  ...... : NA
Créé le ...... : 06/06/2006
Modifié le ... :   /  /    
Description .. : Double click sur la grille 
Mots clefs ... : 
*****************************************************************}
procedure TOF_RESULTSIMUL.GrilleDblClick(Sender: TObject);
begin
AgllanceFiche('PAY','RESULTSIMUL','',Getfield('PSR_CODESIMUL')+';'+Getfield('PSR_SALARIE'),'');
end;

Initialization
  registerclasses ( [ TOF_RESULTSIMUL ] ) ; 
end.
