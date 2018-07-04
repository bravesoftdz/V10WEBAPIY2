{***********UNITE*************************************************
Auteur  ...... : PH
Créé le ...... : 10/09/2001
Modifié le ... :   /  /    
Description .. : Unit de gestion du multicritère gestion des guides des 
Suite ........ : écritures comptables
Mots clefs ... : PAIE;GENERCOMPTA
*****************************************************************}
{
PT1 21/08/2003 PH V_421 FQ 10146 Le modèle CEGID n'est plus modifiable

}
unit UTofPG_GUIDEECRMUL;

interface
uses  Windows,
      StdCtrls,
      Controls,
      Classes,
      Graphics,
      forms,
      sysutils,
      ComCtrls,
      Vierge,
      HCtrls,
      pgOutils,
      UTOF,
      AGLInit ;
Type
     TOF_PG_GUIDEECRMUL = Class (TOF)
       private
       PGPredef, PGNodossier, PGJeu : String;
      public
       procedure OnArgument(Arguments : String ) ; override ;
       procedure OnLoad ; override;
     END ;

implementation


procedure TOF_PG_GUIDEECRMUL.OnArgument(Arguments: String);
var    st : String;
       CEG,STD,DOS : boolean;
begin
  inherited ;
  st:= Trim (Arguments);
  PGPredef:=ReadTokenSt(st);   // Recup Predefini
  PGNodossier:=ReadTokenSt(st);  // Recup Nodossier
  PGJeu:=ReadTokenSt(st);  // Recup modele
// PT1 21/08/2003 PH V_421 FQ 10146 Le modèle CEGID n'est plus modifiable
  AccesPredefini('TOUS',CEG,STD,DOS);
  if (PGPredef = 'CEG')  then
  begin
    PaieLectureSeule(TFVierge(Ecran),(CEG=False));
    SetControlEnabled('BInsert',CEG);
  end
  else
  if (PGPredef = 'STD')  then
  begin
    PaieLectureSeule(TFVierge(Ecran),(STD=False));
    SetControlEnabled('BInsert',STD);
  end
  else
  if (PGPredef = 'DOS')  then
  begin
    PaieLectureSeule(TFVierge(Ecran),False);
    SetControlEnabled('BInsert',DOS);
  end;
SetControlEnabled('BAnnuler',True);
// FIN PT1
end;


procedure TOF_PG_GUIDEECRMUL.OnLoad;
begin
  inherited;
  if PGPredef <> '' then SetControlProperty  ('PGC_PREDEFINI', 'Value', PGPredef);
  if PGNodossier <> '' then SetControlProperty ('PGC_NODOSSIER','Text', PGNodossier);
  if PGJeu <> '' then SetControlProperty ('PGC_JEUECR','Value', PGJeu);
end;

Initialization
registerclasses([TOF_PG_GUIDEECRMUL]);
end.
