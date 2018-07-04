{***********UNITE*************************************************
Auteur  ...... :
Créé le ...... : 17/10/2002
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : MAILINGSTD ()
                 TOF héritable par une autre TOF
Mots clefs ... : TOF;MailingStd
*****************************************************************}
Unit UTOFMailingStd;

Interface

Uses
   Classes, AGLUtilOLE, FileCtrl,
{$IFNDEF EAGLCLIENT}

{$ENDIF}
   UTOF, Menus, HQry, HTB97, UTOFZonesLibres;

Type
   TOF_MAILINGSTD = Class (TOF_ZONESLIBRES)
      private
         pmPopUpMenu_f    : TPopUpMenu;         // Menu PopUp du bouton BMAQUETTE
         btDonnees_f      : TToolbarButton97;   // Bouton BDONNEES
         btFicMailing_f   : TToolbarButton97;   // Bouton BFICFUSION
         btPrintMailing_f : TToolbarButton97;   // Bouton BPRINTFUSION
         QRYFiche_f       : THQuery;            // Query Q de la fiche
         sRepMaquette_f   : string;             // Répertoire des fichiers maquettes
         sRepExport_f     : string;             // Répertoire des fichiers d'export de données
         sRepResultat_f   : string;             // Répertoire des fichiers résultats

         procedure OnClickBtNewMaquette( Sender : TObject );
         procedure OnClickBtOldMaquette( Sender : TObject );
         procedure OnClickBtDonnees( Sender : TObject );
         procedure OnClickBtFicMailing( Sender : TObject );
         procedure OnClickBtPrintMailing( Sender : TObject );
      public
         procedure OnArgument (sParams_p : String ) ; override ;
         procedure RepertoireMailing( sRepMaquette_p : string;
                     sRepResultat_p : string; sRepExport_p : string );
         procedure RepertoireControl( var sRepertoire_p : string; sValeur_p : string );
   end ;

Implementation

Uses
   MailingWord {UtilWord}, DpJurOutils;
     // MailingWord à la place de l'AGL UtilWord qui ne permet pas
     // d'initialiser les répertoire par défaut

{*****************************************************************
Auteur ....... : BM
Date ......... : 17/10/02
Procédure .... : OnClickBtFicMailing
Description .. : Click sur bouton fusion vers fichier
Paramètres ... : L'objet
*****************************************************************}

procedure TOF_MAILINGSTD.OnArgument ( sParams_p : String ) ;
begin
  Inherited ;
   RepertoireMailing( '', '', '' );

   pmPopUpMenu_f    := TPopUpMenu( GetControl( 'POPFUSION' ) );
   btDonnees_f      := TToolbarButton97( GetControl( 'BDONNEES' ) );
   btFicMailing_f   := TToolbarButton97( GetControl( 'BFICFUSION' ) );
   btPrintMailing_f := TToolbarButton97( GetControl( 'BPRINTFUSION' ) );
   QRYFiche_f       := THQuery( GetControl( 'Q' ) );


   pmPopUpMenu_f.Items.Items[0].OnClick := OnClickBtNewMaquette;
   pmPopUpMenu_f.Items.Items[1].OnClick := OnClickBtOldMaquette;
   btDonnees_f.OnClick := OnClickBtDonnees;
   btFicMailing_f.OnClick := OnClickBtFicMailing;
   btPrintMailing_f.OnClick := OnClickBtPrintMailing;

end ;

{*****************************************************************
Auteur ....... : BM
Date ......... : 17/10/02
Procédure .... : RepertoireMailing
Description .. : Initialisation du répertoire des maquettes, des
                 fichiers résultat et des ficheirs d'export
                 Par défaut, répertoire par défaut de WORD
Paramètres ... : Le répertoire des fichiers maquettes
                 Le répertoire des fichiers résultat
                 Le répertoire des fichiers d'export
*****************************************************************}

procedure TOF_MAILINGSTD.RepertoireMailing( sRepMaquette_p  : string;
                     sRepResultat_p : string; sRepExport_p : string );
begin
   RepertoireControl( sRepMaquette_f, sRepMaquette_p );
   RepertoireControl( sRepResultat_f, sRepResultat_p );
   RepertoireControl( sRepExport_f, sRepExport_p );
end;

{***********A.G.L.***********************************************
Auteur  ...... : B.MERIAUX
Créé le ...... : 03/09/2003
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure TOF_MAILINGSTD.RepertoireControl( var sRepertoire_p : string; sValeur_p : string );
begin
   if ( sValeur_p <> '' ) and DirectoryExists( sValeur_p ) then
      sRepertoire_p := sValeur_p
   else
      sRepertoire_p := GetWordTemplateDirectory;
end;
{*****************************************************************
Auteur ....... : BM
Date ......... : 17/10/02
Procédure .... : OnClickBtNewMaquette
Description .. : Click sur bouton nouvelle maquette
Paramètres ... : L'objet
*****************************************************************}

procedure TOF_MAILINGSTD.OnClickBtNewMaquette( Sender : TObject );
begin
   LanceFusionNew( QRYFiche_f, sRepMaquette_f + '\' );
end;

{*****************************************************************
Auteur ....... : BM
Date ......... : 17/10/02
Procédure .... : OnClickBtOldMaquette
Description .. : Click sur bouton ouvrir une maquette
Paramètres ... : L'objet
*****************************************************************}

procedure TOF_MAILINGSTD.OnClickBtOldMaquette( Sender : TObject );
begin
   LanceFusionOpen( QRYFiche_f, sRepMaquette_f + '\' );
end;
{*****************************************************************
Auteur ....... : BM
Date ......... : 17/10/02
Procédure .... : OnClickBtDonnees
Description .. : Click sur bouton export des données
Paramètres ... : L'objet
*****************************************************************}

procedure TOF_MAILINGSTD.OnClickBtDonnees( Sender : TObject );
var
   sFicExport_l : string;
begin
   sFicExport_l := AffecteNomDocMaquette( sRepExport_f  + '\Export.doc' );
   LanceExportDataSource( QRYFiche_f, sFicExport_l );
end;

{*****************************************************************
Auteur ....... : BM
Date ......... : 17/10/02
Procédure .... : OnClickBtFicMailing
Description .. : Click sur bouton fusion vers fichier
Paramètres ... : L'objet
*****************************************************************}

procedure TOF_MAILINGSTD.OnClickBtFicMailing( Sender : TObject );
var
   sFicResultat_l : string;
begin
   sFicResultat_l := AffecteNomDocMaquette( sRepResultat_f  + '\Resultat.doc' );
   LanceFusionWord ( QRYFiche_f, 'FILE', sRepMaquette_f + '\', sFicResultat_l );
end;

{*****************************************************************
Auteur ....... : BM
Date ......... : 17/10/02
Procédure .... : OnClickBtPrintMailing
Description .. : Click sur bouton fusion vers imprimante
Paramètres ... : L'objet
*****************************************************************}

procedure TOF_MAILINGSTD.OnClickBtPrintMailing( Sender : TObject );
begin
   LanceFusionWord (QRYFiche_f, 'PRINTER', sRepMaquette_f + '\' );
end;

Initialization
  registerclasses ( [ TOF_MAILINGSTD ] ) ;
end.

