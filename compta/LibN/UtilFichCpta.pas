unit UtilFichCpta;

interface

Uses StdCtrls,
     Controls,
     Classes,
{$IFNDEF EAGLCLIENT}
     db,
     dbtables,
{$ENDIF}
     forms,
     sysutils,
     ComCtrls,
     HCtrls,
     HEnt1,
     HMsgBox,
     Ent1
      ;


// Sections analytiques
Function EstDansGuideAna(Stc,Sta : String) : Boolean ;
Function EstDansAnalytiq(Stc,Sta : String) : Boolean ;
Function EstDansAxe(Stc,Sta : String) : Boolean ;
Function EstCorresp(Stc,Sta : String) : Boolean ;
Function EstDansVentil(Stc,Sta : String) : Boolean ;
Function ExisteLeCompte(LeWhere : String) : Boolean ;

implementation

{***********A.G.L.***********************************************
Auteur  ...... : Nathalie Payrot
Créé le ...... : 07/01/2002
Modifié le ... :   /  /
Description .. : Fonction qui indique si la section (stc) est dans un guide
Suite ........ : analytique pour l'axe (sta)
Mots clefs ... : SECTION;AXE;GUIDE
*****************************************************************}
Function EstDansGuideAna(Stc,Sta : String) : Boolean ;
BEGIN
  Result := ExisteSql('Select AG_SECTION From ANAGUI Where AG_SECTION="'+Stc+'" AND AG_AXE="'+Sta+'"') ;
END ;

{***********A.G.L.***********************************************
Auteur  ...... : Nathalie Payrot
Créé le ...... : 07/01/2002
Modifié le ... :   /  /
Description .. : Fonction indiquant si une section est utilisée pour une
Suite ........ : écriture analytique.
Mots clefs ... : SECTION;AXE;ANALYTIQUE
*****************************************************************}
Function EstDansAnalytiq(Stc,Sta : String) : Boolean ;
BEGIN
  Result := ExisteSql('Select Y_SECTION From ANALYTIQ Where Y_SECTION="'+Stc+'" AND Y_AXE="'+Sta+'"') ;
END ;

{***********A.G.L.***********************************************
Auteur  ...... : Nathalie Payrot
Créé le ...... : 07/01/2002
Modifié le ... :   /  /
Description .. : Fonction indiquant si la section est section d'attente de
Suite ........ : l'axe.
Mots clefs ... : SECTION;AXE;ATTENTE
*****************************************************************}
Function EstDansAxe(Stc,Sta : String) : Boolean ;
BEGIN
  Result:=(Stc=VH^.Cpta[AxeToFb(Sta)].Attente) ;
END ;

{***********A.G.L.***********************************************
Auteur  ...... : Nathalie Payrot
Créé le ...... : 07/01/2002
Modifié le ... :   /  /
Description .. : Fonction indiquant si la section (stc) est une section de
Suite ........ : correspondance.
Mots clefs ... : SECTION;AXE;CORRESPONDANCE
*****************************************************************}
Function EstCorresp(Stc,Sta : String) : Boolean ;
BEGIN
  Result := ExisteSql('Select S_CORRESP1,S_CORRESP2 From SECTION Where (S_CORRESP1="'+Stc+'" OR S_CORRESP2="'+Stc+'") AND (S_AXE="'+Sta+'")') ;
END ;

{***********A.G.L.***********************************************
Auteur  ...... : Nathalie Payrot
Créé le ...... : 08/01/2002
Modifié le ... :   /  /
Description .. : Fonction indiquant si la section est utilisée dans une
Suite ........ : ventilation type.
Mots clefs ... : SECTION;AXE;VENTILATION
*****************************************************************}
Function EstDansVentil(Stc,Sta : String) : Boolean ;
BEGIN
  Result:=Presence('VENTIL','V_SECTION',Stc) ;
END ;

{***********A.G.L.***********************************************
Auteur  ...... : G.Verdier
Créé le ...... : 16/01/2002
Modifié le ... :   /  /
Description .. : Fonction indiquant si le compte existe dans le plan
Suite ........ : comptable
Mots clefs ... : GENERAL
*****************************************************************}
Function ExisteLeCompte(LeWhere : String) : Boolean ;
BEGIN
Result:=ExisteSql('Select G_GENERAL From GENERAUX '+LeWhere) ;
END ;


end.
