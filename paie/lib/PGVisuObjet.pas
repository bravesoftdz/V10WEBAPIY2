{***********UNITE*************************************************
Auteur  ...... : PH
Créé le ...... : 30/07/2001
Modifié le ... : 30/07/2001
Description .. : Unit qui permet de visualiser un objet de type TOB et de
Suite ........ : l'imprimer à partir d'un grid.
Suite ........ : A utiliser pour le développement pour éviter de faire un
Suite ........ : TOBVIEIWER
Mots clefs ... : PAIE
*****************************************************************}
{
PT1  : 12/08/2002 V585 PH Traitement de la liste des champs à mettre dans la grille
}
unit PGVisuObjet;

interface

uses  Classes,HCtrls,sysutils,HSysMenu,Grids,
{$IFDEF EAGLCLIENT}
      MaineAGL,UtileAgl,
{$ELSE}
      FE_Main,PrintDbg,
{$ENDIF}
      AGLInit,HTB97,UTOF,Utob,Comctrls, Vierge;
//
Type

     TOF_PGVisuObjet = Class (TOF)
     public
        procedure OnLoad ; override ;
        procedure OnArgument(Arguments : String ) ; override ;
     private
        LeTitre,SousTitre,LesChamps : String;
        procedure ImpClik (Sender: TObject);
     END ;

procedure PGVisuUnObjet ( UnObjet : Tob; Titre, SousTitre : String ; PLesCHamps: STRING='');

implementation

procedure PGVisuUnObjet ( UnObjet : Tob; Titre, SousTitre : String ; PLesCHamps: STRING='');
begin
TheTob:=UnObjet ;
AGLLanceFiche('PAY','PGVISUOBJET','','',StringReplace(Titre,';',',',[rfReplaceAll])+';'+SousTitre+';'+StringReplace(PLesChamps,';',',',[rfReplaceAll]));
end ;

procedure TOF_PGVisuObjet.ImpClik (Sender: TObject);
var LeGrid : THGrid;
begin
{$IFNDEF EAGLCLIENT}
LeGrid := THGrid (GetControl ('GRILLE'));
if LeGrid <> NIL then
  PrintDBGrid ( TCustomGrid(LeGrid) , NIL,'Génération des Ods de paie','' ) ;
{$ELSE}
  PrintDBGrid ( 'GRILLE' , '','Génération des Ods de paie','' ) ;
{$ENDIF}
end;


procedure TOF_PGVisuObjet.OnArgument(Arguments: String);
var  // V_42 Qualité
    BtnImp  : TToolbarButton97;
begin
  inherited;
LeTitre   := ReadTokenSt(Arguments);
SousTitre := ReadTokenSt(Arguments);
// PT1  : 12/08/2002 V585 PH Traitement de la liste des champs à mettre dans la grille
LesChamps := ReadTokenSt(Arguments);
if LesChamps <> '' then LesChamps := StringReplace(LesChamps, ',', ';', [rfReplaceAll]);

BtnImp    := TToolbarButton97 (GetControl ('Bimprimer'));
If BtnImp <> NIL then BtnImp.OnClick := ImpClik;
end;

Procedure  TOF_PGVisuObjet.Onload ;
var LeLabel : THLABEL;
//    HMTrad    : THSystemMenu;

begin
inherited ;
// PT1  : 12/08/2002 V585 PH Traitement de la liste des champs à mettre dans la grille
if LesChamps = '' then LaTob.PutGridDetail(THGRID(GetControl('Grille')),True,True,'',True)
 else LaTob.PutGridDetail(THGRID(GetControl('Grille')),True,True,LesChamps,True);

 If (Ecran<>nil) and (LeTitre<>'') then Ecran.Caption:= LeTitre;
LeLabel := THLABEL (GetControl ('LBLSSTITRE'));
if LeLabel <> NIL then LeLabel.Caption := SousTitre;
//HMTrad.ResizeGridColumns (THGRID(GetControl('Grille')));
end;


Initialization
registerclasses([TOF_PGVisuObjet, TUpDown]);
end.
