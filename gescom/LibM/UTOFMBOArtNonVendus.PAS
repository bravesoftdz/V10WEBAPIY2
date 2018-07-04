{***********UNITE*************************************************
Auteur  ...... : Olivier TARCY
Créé le ...... : 15/02/2002
Modifié le ... : 25/02/2002
Description .. : Source TOF de la FICHE : MBOARTNONVENDUS ()
Suite ........ : Edtion des articles non vendus =
Suite ........ :  Stock physique > 0 et Ventes FFO = 0
Mots clefs ... : TOF;MBOARTNONVENDUS
*****************************************************************}
Unit UtofMBOArtNonVendus ;

Interface

Uses StdCtrls,
     Controls, 
     Classes, 
{$IFNDEF EAGLCLIENT}
     db, 
     dbtables,
//     eQRS1,
{$ENDIF}
     forms, 
     sysutils, 
     ComCtrls,
     HCtrls, 
     HEnt1, 
     HMsgBox, 
     UTOF,
     QRS1,
     HQry ; 

Type
  TOF_MBOARTNONVENDUS = Class (TOF)
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
  end ;

Implementation

procedure TOF_MBOARTNONVENDUS.OnLoad ;
var stSQL : string ;
begin
  Inherited ;
  //
  // Initialisation des variable des dates de la période de vente
  //
  SetControlText('XX_VARIABLE1','');
  SetControlText('XX_VARIABLE2','');
  SetControlText('XX_VARIABLE3','');
  SetControlText('XX_VARIABLE1', USDateTime(StrToDate(GetControlText('DATEPIECE1')))) ;
  SetControlText('XX_VARIABLE2', USDateTime(StrToDate(GetControlText('DATEPIECE2')))) ;
  SetControlText('XX_VARIABLE3', GetControlText('FOURNISSEUR')) ;
end ;

procedure TOF_MBOARTNONVENDUS.OnArgument (S : String ) ;
begin
  Inherited ;
  //
  // Affichage du libellé des familles
  //
  THLabel(TForm(Ecran).FindComponent('TGA_FAMILLENIV1')).Caption := RechDom('GCLIBFAMILLE','LF1',False) ;
  THLabel(TForm(Ecran).FindComponent('TGA_FAMILLENIV2')).Caption := RechDom('GCLIBFAMILLE','LF2',False) ;
  THLabel(TForm(Ecran).FindComponent('TGA_FAMILLENIV3')).Caption := RechDom('GCLIBFAMILLE','LF3',False) ;
end ;

Initialization
  registerclasses ( [ TOF_MBOARTNONVENDUS ] ) ;
end.
