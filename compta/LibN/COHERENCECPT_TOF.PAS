{***********UNITE*************************************************
Auteur  ...... :
Créé le ...... : 03/10/2001
Modifié le ... :   /  /
Description .. : Source TOF de la TABLE : COHERENCECPT ()
Mots clefs ... : TOF;COHERENCECPT
*****************************************************************}
Unit COHERENCECPT_TOF ;

Interface

Uses StdCtrls, Controls, Classes, db, forms, sysutils, dbTables, ComCtrls,
     HCtrls, HEnt1, HMsgBox, UTOF, UTob, HTB97, Grids ;
                       
Type
  TOF_COHERENCECPT = Class (TOF)
    TobCompte : TOB ;
    GrilleErreur : THGrid ;
    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnClose                  ; override ;
  private
    procedure BImprimer_OnClick(sender: TObject);
    procedure ControleCoherence;
    procedure RechercheCompte;
    procedure SupprimeCorrect;
    procedure MarquerErreur(NumErr: integer; var TobCpt : Tob);
    procedure BLancerOnClick(sender: TObject);
  end ;

Implementation

Uses PrintDBG ;

procedure TOF_COHERENCECPT.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_COHERENCECPT.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_COHERENCECPT.OnUpdate ;
begin
  Inherited ;
end ;

procedure TOF_COHERENCECPT.OnLoad ;
begin
  Inherited ;
  TobCompte := Tob.Create('Mes comptes', Nil, -1 ) ;
end ;

procedure TOF_COHERENCECPT.OnArgument (S : String ) ;
begin
  Inherited ;

  TToolBarButton97(GetControl('BIMPRIMER')).OnClick := BImprimer_OnClick ;
  TToolBarButton97(GetControl('BLANCER')).OnClick := BLancerOnClick ;

  GrilleErreur := THGrid(GetControl('GRILLEERREUR')) ;
end ;

procedure TOF_COHERENCECPT.OnClose ;
begin
  TobCompte.Free ;
  Inherited ;
end ;

procedure TOF_COHERENCECPT.BLancerOnClick ( sender : TObject ) ;
begin
  // contrôle cohérence des bornes de comptes saisis
  if GetControlText('G_GENERAL') > GetControlText('G_GENERAL_') then
    begin
    PGIBOX(TraduireMemoire('Les bornes des comptes sont incorrectes.'),TraduireMemoire('Contrôle cohérence nature / code compte')) ;
    SetFocusControl('G_GENERAL') ;
    exit ;
    end ;
  If PGIASK(TraduireMemoire('Confirmez-vous le traitement ?'),TraduireMemoire('Contrôle cohérence nature / code compte'))=mrYes then
    BEGIN
    SetControlProperty('RESULTAT','TABVISIBLE', false ) ;
    SetControlVisible('BIMPRIMER', False ) ;
    // lancement du traitement
    ControleCoherence ;
    END ;
end ;

procedure TOF_COHERENCECPT.BImprimer_OnClick ( sender : TObject ) ;
begin
// impression de la grille
PrintDBGrid(GrilleErreur,Nil,'Contrôle cohérence nature / code compte','')
end ;

procedure TOF_COHERENCECPT.RechercheCompte ;
var ReqSql : string ;
    Where : string ;
    Q : TQuery ;
begin
  // recherche des comptes concernés
  ReqSql := 'SELECT G_GENERAL, G_LIBELLE, G_NATUREGENE FROM GENERAUX' ;
  Where := '' ;
  If GetControlText('G_GENERAL') <> '' then Where := ' WHERE G_GENERAL>="' +GetControlText('G_GENERAL')+'"' ;
  If GetControlText('G_GENERAL_') <> '' then
    begin
      If Where = '' then Where := ' WHERE'
      Else Where := Where + ' AND' ;
      Where := Where + ' G_GENERAL<="' +GetControlText('G_GENERAL_')+'"' ;
    end ;
  Q := OpenSql(ReqSql+Where, True ) ;
  TobCompte.LoadDetailDB('Mes comptes', '', '', Q, False, True ) ;
  Ferme(Q) ;
end ;

procedure TOF_COHERENCECPT.SupprimeCorrect ;
Var i : integer ;
    TobC  : TOB ;
begin
// Suppression des comptes corrects de la TOB
for i:=TobCompte.Detail.Count-1 downto 0 do
  begin
  TobC:=TobCompte.Detail[i] ;
  if TobC.GetValue('ERREUR')<>'X' then begin TobC.Free ; TobC:=Nil ; end ;
  end ;
end ;

procedure TOF_COHERENCECPT.MarquerErreur ( NumErr : integer; var TobCpt : Tob ) ;
var Mess : string ;
begin
  // maj des comptes incorrects dans la tob avec le message adéquat
  Case NumErr of
  0 : Mess := TraduireMemoire('Le compte commence par 0.') ;
  1 : Mess := TraduireMemoire('Les comptes de nature Banque doivent être de 512 à 5149.') ;
  2 : Mess := TraduireMemoire('Le compte devrait avoir une nature Banque.') ;
  3 : Mess := TraduireMemoire('Les comptes de nature Banque doivent être de 530 à 531.') ;
  4 : Mess := TraduireMemoire('Le compte devrait avoir une nature Caisse.') ;
  5 : Mess := TraduireMemoire('La nature Charge correspond aux comptes de la classe 6.') ;
  6 : Mess := TraduireMemoire('Le compte devrait avoir une nature Charge.') ;
  7 : Mess := TraduireMemoire('La nature Produit correspond aux comptes de la classe 7.') ;
  8 : Mess := TraduireMemoire('Le compte devrait avoir une nature Produit.') ;
  9 : Mess := TraduireMemoire('Les comptes de nature Collectifs Client doivent être 41 ou 491.') ;
  10 : Mess := TraduireMemoire('Le compte devrait avoir une nature Collectif Client ou Divers.') ;
  11 : Mess := TraduireMemoire('Les comptes de nature Collectifs Client doivent être 40 ou 490.') ;
  12 : Mess := TraduireMemoire('Le compte devrait avoir une nature Collectif Fournisseur ou Divers.') ;
  13 : Mess := TraduireMemoire('Les comptes de nature Collectifs Salarié doivent être 421.') ;
  14 : Mess := TraduireMemoire('Le compte devrait avoir une nature Collectif Salarié ou Divers.') ;
  15 : Mess := TraduireMemoire('La nature Immobilisation correspond aux comptes de la classe 2.') ;
  16 : Mess := TraduireMemoire('Le compte devrait avoir une nature Immobilisation ou Collectif Divers.') ;
  17 : Mess := TraduireMemoire('Les comptes de classe 8 ou 9 doivent être de nature Divers ou Extra-Comptable.') ;
  18 : Mess := TraduireMemoire('Le compte commence par un caractère non numérique.') ;
  end ;
  TobCpt.PutValue('ERREUR', 'X' ) ;
  TobCpt.PutValue('LIBERREUR', Mess) ;
end ;

procedure TOF_COHERENCECPT.ControleCoherence ;
var IdxCpte : integer ;
    TC : TOB ;
begin
  EnableControls(Ecran, false ) ;
  // remise à 0 de la TOB et de la grille
  TobCompte.ClearDetail ;
  GrilleErreur.VidePile(False ) ;
  // recherche des comptes
  RechercheCompte ;
  If TobCompte.Detail.Count > 0 then
    begin
    // ajout des colonnes Erreur et libellé de l'erreur à la TOB
    TobCompte.Detail[0].AddChampSupValeur('ERREUR','',True ) ;
    TobCompte.Detail[0].AddChampSupValeur('LIBERREUR','',True ) ;
    // pour chaque compte trouvé, on contrôle la cohérence entre la nature et le code
    For idxCpte := 0 to TobCompte.Detail.Count-1 do
      begin
      TC := TobCompte.Detail[IdxCpte] ;
      // Compte commence par 0
      If Copy(TC.GetValue('G_GENERAL'),1,1)='0' then MarquerErreur(0,TC)
      else
       if (Copy(TC.GetValue('G_GENERAL'),1,1)<'0') or (Copy(TC.GetValue('G_GENERAL'),1,1)>'9') then MarquerErreur(18,TC)
       else
        begin
          // Compte de nature banque avec un code non compris entre 512 et 5149
          If (TC.GetValue('G_NATUREGENE')='BQE')
             And ((Copy(TC.GetValue('G_GENERAL'),1,4)< '5120')or(Copy(TC.GetValue('G_GENERAL'),1,4)>'5149')) then MarquerErreur(1,TC) ;
          If ((Copy(TC.GetValue('G_GENERAL'),1,4)>= '5120')and(Copy(TC.GetValue('G_GENERAL'),1,4)<='5149'))
            And (TC.GetValue('G_NATUREGENE')<>'BQE') then MarquerErreur(2,TC) ;

          // Compte de nature caisse avec un code non compris entre 530 et 531
          If (TC.GetValue('G_NATUREGENE')='CAI')
             And ((Copy(TC.GetValue('G_GENERAL'),1,3)< '530')or(Copy(TC.GetValue('G_GENERAL'),1,3)>'531')) then MarquerErreur(3,TC) ;
          If ((Copy(TC.GetValue('G_GENERAL'),1,3)>= '530')and(Copy(TC.GetValue('G_GENERAL'),1,3)<='531'))
            And (TC.GetValue('G_NATUREGENE')<>'CAI') then MarquerErreur(4,TC) ;

          // Compte de charge ne commençant pas par 6
          If (TC.GetValue('G_NATUREGENE')='CHA')
             And (Copy(TC.GetValue('G_GENERAL'),1,1)<> '6') then MarquerErreur(5,TC) ;
          If (Copy(TC.GetValue('G_GENERAL'),1,1)= '6')
            And (TC.GetValue('G_NATUREGENE')<>'CHA') then MarquerErreur(6,TC) ;

          // Compte de produit ne commençant pas par 7
          If (TC.GetValue('G_NATUREGENE')='PRO')
             And (Copy(TC.GetValue('G_GENERAL'),1,1)<> '7') then MarquerErreur(7,TC) ;
          If (Copy(TC.GetValue('G_GENERAL'),1,1)= '7')
            And (TC.GetValue('G_NATUREGENE')<>'PRO') then MarquerErreur(8,TC) ;

          // Compte collectif client ne commençant pas par 41
          If (TC.GetValue('G_NATUREGENE')='COC')
             and ((Copy(TC.GetValue('G_GENERAL'),1,2)<>'41')
             and(Copy(TC.GetValue('G_GENERAL'),1,3)<>'491')) then MarquerErreur(9,TC) ;
          If ((Copy(TC.GetValue('G_GENERAL'),1,2)= '41')or (Copy(TC.GetValue('G_GENERAL'),1,3)='491'))
            And ((TC.GetValue('G_NATUREGENE')<>'COC')and(TC.GetValue('G_NATUREGENE')<>'DIV')) then MarquerErreur(10,TC) ;

          // Compte collectif fournisseur ne commençant pas par 40
          If (TC.GetValue('G_NATUREGENE')='COF')
             And ((Copy(TC.GetValue('G_GENERAL'),1,2)<> '40')
             and (Copy(TC.GetValue('G_GENERAL'),1,3)<>'490')) then MarquerErreur(11,TC) ;
          If ((Copy(TC.GetValue('G_GENERAL'),1,2)= '40')or(Copy(TC.GetValue('G_GENERAL'),1,3)='490'))
            And ((TC.GetValue('G_NATUREGENE')<>'COF')and(TC.GetValue('G_NATUREGENE')<>'DIV')) then MarquerErreur(12,TC) ;

          // Compte collectif salarié ne commençant pas par 421
          If (TC.GetValue('G_NATUREGENE')='COS')
             And (Copy(TC.GetValue('G_GENERAL'),1,3)<> '421') then MarquerErreur(13,TC) ;
          If (Copy(TC.GetValue('G_GENERAL'),1,3)= '421')
            And ((TC.GetValue('G_NATUREGENE')<>'COS')and(TC.GetValue('G_NATUREGENE')<>'DIV')) then MarquerErreur(14,TC) ;

          // Compte d'imobilisation ne commençant pas par 2
          If (TC.GetValue('G_NATUREGENE')='IMO')
             And (Copy(TC.GetValue('G_GENERAL'),1,1)<> '2') then MarquerErreur(15,TC) ;
          If (Copy(TC.GetValue('G_GENERAL'),1,1)= '2')
            And (TC.GetValue('G_NATUREGENE')<>'IMO')and(TC.GetValue('G_NATUREGENE')<>'COD') then MarquerErreur(16,TC) ;

          // Compte Divers ou Extr-comptable ne commençant pas par 8 ou 9
          If ((Copy(TC.GetValue('G_GENERAL'),1,1)= '8')or (Copy(TC.GetValue('G_GENERAL'),1,1)= '9'))
            And ((TC.GetValue('G_NATUREGENE')<>'DIV')and(TC.GetValue('G_NATUREGENE')<>'EXT')) then MarquerErreur(17,TC) ;
        end ;
      end ;
    // suppression des comptes corrects de la TOB
    SupprimeCorrect ;

    EnableControls(Ecran, true ) ;

    // S'il y a des incohérences
    If TobCompte.Detail.Count > 0 then
      begin
      PGIBox(TraduireMemoire('Des incohérences dans les comptes ont été trouvées.'),TraduireMemoire('Contrôle cohérence nature / code compte')) ;
      // affichage de l'onglet Resultat et bouton Imprimer
      SetControlProperty('RESULTAT','TABVISIBLE', True ) ;
      SetActiveTabSheet('RESULTAT') ;
      SetControlVisible('BIMPRIMER', true) ;
      // affichage de la TOB dans la grille
      TobCompte.PutGridDetail(GrilleErreur, false, false, 'G_GENERAL;G_LIBELLE;G_NATUREGENE;LIBERREUR' ) ;
      end
    else PGIBox(TraduireMemoire('Traitement terminé. Aucune incohérence trouvée.'),TraduireMemoire('Contrôle cohérence nature / code compte')) ;
    end
  else PGIBox(TraduireMemoire('Aucun compte à contrôler.'),TraduireMemoire('Contrôle cohérence nature / code compte')) ;


end ;

Initialization
  registerclasses ( [ TOF_COHERENCECPT ] ) ;
end.
