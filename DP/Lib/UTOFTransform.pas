{***********UNITE*************************************************
Auteur  ...... : 
Créé le ...... : 25/07/2002
Modifié le ... : 06/06/2003
Description .. : Source TOF de la FICHE : TRANSFORM ()
Mots clefs ... : TOF;TRANSFORM
*****************************************************************}
Unit UTOFTransform;
//////////////////////////////////////////////////////////////////
Interface
//////////////////////////////////////////////////////////////////
Uses
   UTOB, UTOF, HCtrls, HTB97;
//////////////////////////////////////////////////////////////////
Type
     TOF_TRANSFORM = Class (TOF)
       procedure OnLoad                   ; override ;
       procedure OnArgument ( sParams_p : String ) ; override ;
       procedure OnClose                  ; override ;

       private
         // Paramètres de la fiche
         sGuidPerDos_fp : string;
         sNomPer_fp : string;
         sAncFormeCode_fp : string;
         sNouFormeCode_fp : string;

         // Variables
         bModif_f : boolean;
         TOBLiensTitre_f : TOB;
         TOBLiensDirig_f : TOB;
         TOBLiensAutre_f : TOB;
         TOBFonctTitre_f : TOB;
         TOBFonctDirig_f : TOB;
         TOBAncForme_f : TOB;
         TOBNouForme_f : TOB;

         // Composants de la fiche
         edAncForme_f : THEdit;
         edNouForme_f : THEdit;

         // Index de colonnes des anciens liens
         nLAColNom_f : integer;
         nLAColCod_f : integer;
         nLAColFon_f : integer;
         nLAColInf_f : integer;
         nLAColToL_f : integer;
         nLAColTri_f : integer;

         // Index de colonnes des nouveaux liens
         nLNColNom_f : integer;
         nLNColCod_f : integer;
         nLNColFon_f : integer;
         nLNColInf_f : integer;
         nLNColToL_f : integer;
         nLNColToF_f : integer;
         nLNColTri_f : integer;

         // Index de colonnes des fonctions
         nFOColFon_f : integer;
         nFOColToF_f : integer;

         fgLiensTitre_f : THGrid;
         fgLiensDirig_f : THGrid;
         fgLiensAutre_f : THGrid;
         fgFonctTitre_f : THGrid;
         fgFonctDirig_f : THGrid;
         fgLiensTitreNew_f : THGrid;
         fgLiensDirigNew_f : THGrid;
         fgLiensAutreNew_f : THGrid;

         btSupTitre_f: TToolbarButton97;
         btChangeTitre_f: TToolbarButton97;
         btAnnulTitre_f: TToolbarButton97;
         btSupDirig_f: TToolbarButton97;
         btChangeDirig_f: TToolbarButton97;
         btAnnulDirig_f: TToolbarButton97;
         btSupAutre_f: TToolbarButton97;
         btConserveAutre_f: TToolbarButton97;
         btAnnulAutre_f: TToolbarButton97;
         btValider_f: TToolbarButton97;
         btFermer_f: TToolbarButton97;

         procedure AnnulerOnClick(Sender : TObject);
         procedure ChangerOnClick(Sender : TObject);
         procedure FermerOnClick(Sender : TObject);
         procedure LigneOnSelect(Sender : TObject);
         procedure LigneNewOnSelect(Sender : TObject);
         procedure SupprimerOnClick(Sender : TObject);
         procedure ValiderOnClick(Sender : TObject);

         procedure AfficherFonctions( TOBLiens_p : TOB; var fgFonction_p : THGrid );
         procedure AfficherLiens( TOBLiens_p : TOB; var fgLiens_p : THGrid );
         function  AjouteLigne( var fgGrille_p : THGrid; nRow_p : integer ) : integer;
         procedure Annuler( TOBLiens_p : TOB; var fgLiens_p : THGrid; var fgLiensNew_p : THGrid );
         procedure Changer( var fgLiens_p : THGrid; fgFonction_p : THGrid; var fgLiensNew_p : THGrid; TOBFonctions_p : TOB );
         function  ChargerFonctions( sWhere_p, sNom_p : string ) : TOB;
         function  ChargerForme( sCodeForme_p : string ) : TOB;
         function  ChargerLiens( sWhere_p, sNom_p : string ) : TOB;
         procedure ColGetLargeur( fgGrille_p : THGrid; var tnColLargeur_p : array of integer );
         procedure ColNewLargeur( fgGrille_p : THGrid; nIndice_p : integer; var tnColLargeur_p : array of integer );
         procedure ColSetLargeur( fgGrille_p : THGrid; var tnColLargeur_p : array of integer );
         procedure EcranVersClasse( fgLiensNew_p : THGrid; var TOBLiens_p : TOB );
         procedure EnleveLigne( var fgGrille_p : THGrid; nRow_p : integer );
         function  Enregistrable : boolean;
         procedure Enregistrer;
         procedure InitGrilles( );
         procedure InitColTitres( var fgGrille_p : THGrid; nColNb_p : integer; sTitres_p : string );
         procedure InitColCache( var fgGrille_p : THGrid; sColCache_p : string );
         procedure MiseAJour( TOBLiens_p, TOBFonctions_p : TOB );
         procedure Supprimer( var fgLiens_p : THGrid );
         procedure ValiderBoutons( sNom_p : string; nLiensRow_p, nLiensFixedRow_p, nLiensRowCount_p : integer; sLiensVal_p : string; nFonctRow_p, nFonctFixedRow_p, nFonctRowCount_p : integer; sFonctVal_p : string );
         function  ValiderLeBouton( sNomBouton_p : string; bForce_p : boolean; nRow_p, nRowFixed, nRowCount_p : integer; sTobVal_p : string ) : boolean;
         function  VerifLien( sNomPer_p, sCodeFonction_p : string; fgLiensNew_p : THGrid ) : boolean;
  end ;
//////////////////////////////////////////////////////////////////
Implementation
//////////////////////////////////////////////////////////////////
Uses
   {$IFNDEF EAGLCLIENT}
   {$IFNDEF DBXPRESS}dbtables,{$ELSE}uDbxDataSet,{$ENDIF}
   {$ENDIF}
   Classes,

   {$IFDEF VER150}
   Variants,
   {$ENDIF}

   sysutils, HMsgBox, Vierge;


//////////////////////////////////////////////////////////////////
{*****************************************************************
Auteur ....... : BM
Date ......... : 25/07/02
Procédure .... : OnArgument
Description .. : Chargement de la fiche
Paramètres ... : La liste des arguments
*****************************************************************}

procedure TOF_TRANSFORM.OnArgument ( sParams_p : String ) ;
begin
   Inherited ;
   sGuidPerDos_fp := ReadTokenSt(sParams_p);
   sNomPer_fp := ReadTokenSt(sParams_p);
   //sGuidPerDos_fp + ' ' +
   Ecran.Caption := Ecran.Caption + ' : ' + sNomPer_fp;
   sAncFormeCode_fp := ReadTokenSt(sParams_p);
   sNouFormeCode_fp := ReadTokenSt(sParams_p);

   edAncForme_f := THEdit(GetControl('OLD_FORME'));
   edNouForme_f := THEdit(GetControl('NEW_FORME'));

   bModif_f := false;

   fgLiensTitre_f := THGrid(GetControl('LIENSTITRE'));
   fgLiensDirig_f := THGrid(GetControl('LIENSDIRIG'));
   fgLiensAutre_f := THGrid(GetControl('LIENSAUTRE'));
   fgFonctTitre_f := THGrid(GetControl('FONCT_TITRE'));
   fgFonctDirig_f := THGrid(GetControl('FONCT_DIRIG'));
   fgLiensTitreNew_f := THGrid(GetControl('LIENSTITRENEW'));
   fgLiensDirigNew_f := THGrid(GetControl('LIENSDIRIGNEW'));
   fgLiensAutreNew_f := THGrid(GetControl('LIENSAUTRENEW'));
   fgLiensTitre_f.OnClick := LigneOnSelect;
   fgLiensDirig_f.OnClick := LigneOnSelect;
   fgLiensAutre_f.OnClick := LigneOnSelect;
   fgFonctTitre_f.OnClick := LigneOnSelect;
   fgFonctDirig_f.OnClick := LigneOnSelect;
   fgLiensTitreNew_f.OnClick := LigneNewOnSelect;
   fgLiensDirigNew_f.OnClick := LigneNewOnSelect;
   fgLiensAutreNew_f.OnClick := LigneNewOnSelect;
   InitGrilles;

   btSupTitre_f := TToolbarButton97(GetControl('B_SUPTITRE'));
   btSupTitre_f.OnClick := SupprimerOnClick;
   btChangeTitre_f := TToolbarButton97(GetControl('B_CHANGETITRE'));
   btChangeTitre_f.OnClick := ChangerOnClick;
   btAnnulTitre_f := TToolbarButton97(GetControl('B_ANNULTITRE'));
   btAnnulTitre_f.OnClick := AnnulerOnClick;

   btSupDirig_f := TToolbarButton97(GetControl('B_SUPDIRIG'));
   btSupDirig_f.OnClick := SupprimerOnClick;
   btChangeDirig_f := TToolbarButton97(GetControl('B_CHANGEDIRIG'));
   btChangeDirig_f.OnClick := ChangerOnClick;
   btAnnulDirig_f := TToolbarButton97(GetControl('B_ANNULDIRIG'));
   btAnnulDirig_f.OnClick := AnnulerOnClick;

   btSupAutre_f := TToolbarButton97(GetControl('B_SUPAUTRE'));
   btSupAutre_f.OnClick := SupprimerOnClick;
   btConserveAutre_f := TToolbarButton97(GetControl('B_CHANGEAUTRE'));
   btConserveAutre_f.OnClick := ChangerOnClick;
   btAnnulAutre_f := TToolbarButton97(GetControl('B_ANNULAUTRE'));
   btAnnulAutre_f.OnClick := AnnulerOnClick;

   btValider_f := TToolbarButton97(GetControl('BVALIDER'));
   btValider_f.OnClick := ValiderOnClick;
   btFermer_f := TToolbarButton97(GetControl('BFERME'));
   btFermer_f.OnClick := FermerOnClick;
end ;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : 
Créé le ...... : 06/06/2003
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... :
*****************************************************************}
procedure TOF_TRANSFORM.OnLoad ;
begin
   Inherited ;

   TOBAncForme_f := ChargerForme( sAncFormeCode_fp );
   TOBNouForme_f := ChargerForme( sNouFormeCode_fp );

   edAncForme_f.Text := TOBAncForme_f.GetValue('JFJ_FORMEABREGE');
   edNouForme_f.Text := TOBNouForme_f.GetValue('JFJ_FORMEABREGE');

   TOBLiensTitre_f := ChargerLiens( 'and ANL_FORME <> "" and JTF_TITRE = "X" ', 'LIENSTITRE' );
   TOBLiensDirig_f := ChargerLiens( 'and ANL_FORME <> "" and JTF_DIRIGEANT = "X" ', 'LIENSDIRIG' );
//   TOBLiensAutre_f := ChargerLiens( 'and ANL_FORME = "" and JTF_TITRE <> "X" and JTF_DIRIGEANT <> "X" and ( ANL_FONCTION not in ("STE","INT") )', 'LIENSAUTRE' );
   TOBLiensAutre_f := ChargerLiens( 'and ANL_FORME = "" and JTF_TITRE <> "X" and JTF_DIRIGEANT <> "X" and ANL_TIERS <> "X"', 'LIENSAUTRE' );
   TOBFonctTitre_f := ChargerFonctions( 'and JTF_TITRE = "X" ', 'FONCTTITRE' );
   TOBFonctDirig_f := ChargerFonctions( 'and JTF_DIRIGEANT = "X" ', 'FONCTDIRIG' );

   AfficherLiens( TOBLiensTitre_f, fgLiensTitre_f );
   AfficherLiens( TOBLiensDirig_f, fgLiensDirig_f );
   AfficherLiens( TOBLiensAutre_f, fgLiensAutre_f );
   AfficherFonctions( TOBFonctTitre_f, fgFonctTitre_f );
   AfficherFonctions( TOBFonctDirig_f, fgFonctDirig_f );

   btValider_f.Enabled := Enregistrable;
      
   LigneOnSelect( fgLiensTitre_f);
   LigneOnSelect( fgLiensDirig_f);
   LigneOnSelect( fgLiensAutre_f);
   LigneNewOnSelect( fgLiensTitreNew_f);
   LigneNewOnSelect( fgLiensDirigNew_f);
   LigneNewOnSelect( fgLiensAutreNew_f);

end ;


{*****************************************************************
Auteur ....... : BM
Date ......... : 25/07/02
Procédure .... : OnClose
Description .. : Fermeture de la fenêtre
Paramètres ... :
*****************************************************************}

procedure TOF_TRANSFORM.OnClose ;
begin
   Inherited ;
   TOBLiensTitre_f.Free;
   TOBLiensDirig_f.Free;
   TOBLiensAutre_f.Free;
   TOBFonctTitre_f.Free;
   TOBFonctDirig_f.Free;
   TOBAncForme_f.Free;
   TOBNouForme_f.Free;


   if bModif_f then
      TFVierge(Ecran).Retour := '1'
   else
      TFVierge(Ecran).Retour := '0';

end ;

{*****************************************************************
Auteur ....... : BM
Date ......... : 25/07/02
Procédure .... : btValiderOnClick
Description .. : Click sur le bouton VALIDER
Paramètres ... : L'objet
*****************************************************************}

procedure TOF_TRANSFORM.ValiderOnClick(Sender : TObject);
begin
   bModif_f := true;
   Enregistrer;
end;

{*****************************************************************
Auteur ....... : BM
Date ......... : 25/07/02
Procédure .... : FermerOnClick
Description .. : Click sur le bouton ANNULER
Paramètres ... : L'objet
*****************************************************************}

procedure TOF_TRANSFORM.FermerOnClick(Sender : TObject);
begin
   bModif_f := false;
end;

{*****************************************************************
Auteur ....... : BM
Date ......... : 26/07/02
Procédure .... : LigneOnSelect
Description .. : Contrôle des boutons SUPPRIMER et CHANGER après la sélection d'une ligne de la grille
Paramètres ... : La grille de liens actuels
*****************************************************************}

procedure TOF_TRANSFORM.LigneOnSelect( Sender : TObject);
begin
   if ( Sender as THGrid ).Name = 'LIENSTITRE' then
   begin
      ValiderBoutons( 'TITRE',
                  fgLiensTitre_f.Row, fgLiensTitre_f.FixedRows, fgLiensTitre_f.RowCount, fgLiensTitre_f.Cells[0,fgLiensTitre_f.Row],
                  fgFonctTitre_f.Row, fgFonctTitre_f.FixedRows, fgFonctTitre_f.RowCount, fgFonctTitre_f.Cells[0,fgFonctTitre_f.Row] );
   end;
   if ( Sender as THGrid ).Name = 'LIENSDIRIG' then
   begin
      ValiderBoutons( 'DIRIG',
                  fgLiensDirig_f.Row, fgLiensDirig_f.FixedRows, fgLiensDirig_f.RowCount, fgLiensDirig_f.Cells[0,fgLiensDirig_f.Row],
                  fgFonctDirig_f.Row, fgFonctDirig_f.FixedRows, fgFonctDirig_f.RowCount, fgFonctDirig_f.Cells[0,fgFonctDirig_f.Row] );
   end;
   if ( Sender as THGrid ).Name = 'LIENSAUTRE' then
   begin
      ValiderBoutons( 'AUTRE',
                  fgLiensAutre_f.Row, fgLiensAutre_f.FixedRows, fgLiensAutre_f.RowCount, fgLiensAutre_f.Cells[0,fgLiensDirig_f.Row],
                  -1, -1, -1, 'TOB' );
   end;
end;

{*****************************************************************
Auteur ....... : BM
Date ......... : 29/07/02
Procédure .... : LigneOnSelect
Description .. : Contrôle des boutons ANNULER après la sélection d'une ligne de la grille
Paramètres ... : La grille de nouveaux liens
*****************************************************************}

procedure TOF_TRANSFORM.LigneNewOnSelect( Sender : TObject);
begin
   if ( Sender as THGrid ).Name = 'LIENSTITRENEW' then
   begin
      btAnnulTitre_f.Enabled := ValiderLeBouton( 'B_ANNULTITRE', true,
                  fgLiensTitreNew_f.Row, fgLiensTitreNew_f.FixedRows,
                  fgLiensTitreNew_f.RowCount, fgLiensTitreNew_f.Cells[0,fgLiensTitreNew_f.Row] );
   end;
   if ( Sender as THGrid ).Name = 'LIENSDIRIGNEW' then
   begin
      btAnnulTitre_f.Enabled := ValiderLeBouton( 'B_ANNULDIRIG', true,
                  fgLiensDirigNew_f.Row, fgLiensDirigNew_f.FixedRows,
                  fgLiensDirigNew_f.RowCount, fgLiensDirigNew_f.Cells[0,fgLiensDirigNew_f.Row] );
   end;
   if ( Sender as THGrid ).Name = 'LIENSAUTRENEW' then
   begin
      btAnnulTitre_f.Enabled := ValiderLeBouton( 'B_ANNULAUTRE', true,
                  fgLiensAutreNew_f.Row, fgLiensAutreNew_f.FixedRows,
                  fgLiensAutreNew_f.RowCount, fgLiensAutreNew_f.Cells[0,fgLiensDirigNew_f.Row] );
   end;

end;

{*****************************************************************
Auteur ....... : BM
Date ......... : 26/07/02
Procédure .... : SupprimerOnClick
Description .. : Click sur le bouton SUPPRIMER des grilles
Paramètres ... : L'objet
*****************************************************************}

procedure TOF_TRANSFORM.SupprimerOnClick(Sender : TObject);
begin
   if ( Sender as TToolBarButton97 ).Name = 'B_SUPTITRE' then
   begin
      Supprimer( fgLiensTitre_f );
   end;
   if ( Sender as TToolBarButton97 ).Name = 'B_SUPDIRIG' then
   begin
      Supprimer( fgLiensDirig_f );
   end;
   if ( Sender as TToolBarButton97 ).Name = 'B_SUPAUTRE' then
   begin
      Supprimer( fgLiensAutre_f );
   end;
   btValider_f.Enabled := Enregistrable;
end;

{*****************************************************************
Auteur ....... : BM
Date ......... : 26/07/02
Procédure .... : ChangerOnClick
Description .. : Click sur le bouton CHANGER ou CONSERVER des grilles
Paramètres ... : L'objet
*****************************************************************}

procedure TOF_TRANSFORM.ChangerOnClick(Sender : TObject);
begin
   if ( Sender as TToolBarButton97 ).Name = 'B_CHANGETITRE' then
   begin
      Changer(fgLiensTitre_f, fgFonctTitre_f, fgLiensTitreNew_f, TOBFonctTitre_f );
   end;
   if ( Sender as TToolBarButton97 ).Name = 'B_CHANGEDIRIG' then
   begin
      Changer(fgLiensDirig_f, fgFonctDirig_f, fgLiensDirigNew_f, TOBFonctDirig_f );
   end;
   if ( Sender as TToolBarButton97 ).Name = 'B_CHANGEAUTRE' then
   begin
      Changer(fgLiensAutre_f, nil, fgLiensAutreNew_f, nil );
   end;
   btValider_f.Enabled := Enregistrable;
end;

{*****************************************************************
Auteur ....... : BM
Date ......... : 26/07/02
Procédure .... : AnnulerOnClick
Description .. : Click sur le bouton ANNULER des grilles
Paramètres ... : L'objet
*****************************************************************}

procedure TOF_TRANSFORM.AnnulerOnClick(Sender : TObject);
begin
   if (Sender as TToolBarButton97 ).Name = 'B_ANNULTITRE' then
   begin
      Annuler( TOBLiensTitre_f, fgLiensTitre_f, fgLiensTitreNew_f );
   end;
   if (Sender as TToolBarButton97 ).Name = 'B_ANNULDIRIG' then
   begin
      Annuler( TOBLiensDirig_f, fgLiensDirig_f, fgLiensDirigNew_f );
   end;
   if (Sender as TToolBarButton97 ).Name = 'B_ANNULAUTRE' then
   begin
      Annuler( TOBLiensAutre_f, fgLiensAutre_f, fgLiensAutreNew_f );
   end;
   btValider_f.Enabled := Enregistrable;
end;

{*****************************************************************
Auteur ....... : BM
Date ......... : 25/07/02
Fonction ..... : ChargerLiens
Description .. : Charge la liste des liens selon la clause where passée en paramêtre
                 et créé la TOB associée
Paramètres ... : Clause where
Renvoie ...... : La TOB
*****************************************************************}

function TOF_TRANSFORM.ChargerLiens( sWhere_p, sNom_p : string ) : TOB;
var
   TOB_l : TOB;
   sRequete_l : string;
begin
   TOB_l := TOB.Create( sNom_p, nil, -1 );
   sRequete_l := 'select ANL_GUIDPER, ANL_NOMPER, ANL_FONCTION, ANL_AFFICHE, ANL_INFO, ANL_TRI ' +
            'from   ANNULIEN, JUTYPEFONCT ' +
            'where  ANL_GUIDPERDOS = "' + sGuidPerDos_fp + '" ' +
            '  and  ANL_TYPEDOS = "STE" ' +
            '  and  ANL_FONCTION = JTF_FONCTION ' +
            sWhere_p +
            'order by ANL_TRI, ANL_NOMPER';
   TOB_l.LoadDetailFromSQL(sRequete_l);
   result := TOB_l;
end;

{*****************************************************************
Auteur ....... : BM
Date ......... : 25/07/02
Fonction ..... : ChargerFonctions
Description .. : Charge la liste des fonctions selon la clause where passée en paramêtre
                 et créé la TOB associée
Paramètres ... : Clause where
Renvoie ...... : La TOB
*****************************************************************}

function TOF_TRANSFORM.ChargerFonctions( sWhere_p, sNom_p : string ) : TOB;
var
   TOB_l : TOB;
   sRequete_l : string;
begin
   TOB_l := TOB.Create( sNom_p, nil, -1 );
   sRequete_l := 'select JFT_FONCTION, JTF_FONCTABREGE, JFT_DEFAUT, JFT_RACINE, JTF_AFFICHE, JFT_TRI ' +
                 'from   JUFONCTION, JUTYPEFONCT ' +
                 'where  JFT_FORME = "' + sNouFormeCode_fp + '" ' +
                 '  and  JFT_FONCTION = JTF_FONCTION ' +
                 sWhere_p +
                 'order by  JFT_TRI';
   TOB_l.LoadDetailFromSQL(sRequete_l);
   result := TOB_l;
end;

{*****************************************************************
Auteur ....... : BM
Date ......... : 25/07/02
Fonction ..... : ChargerForme
Description .. : Charge la forme selon la clause where passée en paramêtre
                 et créé la TOB associée
Paramètres ... : Le code de la forme
Renvoie ...... : La TOB
*****************************************************************}

function TOF_TRANSFORM.ChargerForme( sCodeForme_p : string ) : TOB;
var
   TOB_l : TOB;
begin
   TOB_l := TOB.Create( 'JUFORMEJUR', nil, -1 );
   TOB_l.SelectDB('"' + sCodeForme_p + '"', nil);
   result := TOB_l;
end;

{*****************************************************************
Auteur ....... : BM
Date ......... : 01/08/02
Procédure .... : InitGrilles
Description .. : Initialise le titres et les indexs des colonnes pour les grilles
Paramètres ... :
*****************************************************************}

procedure TOF_TRANSFORM.InitGrilles( );
begin
   nLAColToL_f := 0;
   nLAColNom_f := 1;
   nLAColCod_f := 2;
   nLAColFon_f := 3;
   nLAColInf_f := 4;
   nLAColTri_f := 5;

   InitColTitres( fgLiensTitre_f, 6, 'TOB;Nom;Code;Fonction;Infos;Tri' );
   InitColTitres( fgLiensDirig_f, 6, 'TOB;Nom;Code;Fonction;Infos;Tri' );
   InitColTitres( fgLiensAutre_f, 6, 'TOB;Nom;Code;Fonction;Infos;Tri' );
   InitColCache( fgLiensTitre_f, IntToStr( nLAColToL_f ) + ';' + IntToStr( nLAColTri_f ));
   InitColCache( fgLiensDirig_f, IntToStr( nLAColToL_f ) + ';' + IntToStr( nLAColTri_f ));
   InitColCache( fgLiensAutre_f, IntToStr( nLAColToL_f ) + ';' + IntToStr( nLAColTri_f ));

   nLNColToL_f := 0;
   nLNColNom_f := 1;
   nLNColToF_f := 2;
   nLNColCod_f := 3;
   nLNColFon_f := 4;
   nLNColInf_f := 5;
   nLNColTri_f := 6;

   InitColTitres( fgLiensTitreNew_f, 7, 'TOB;Nom;TOB;Code;Fonction;Infos;Tri' );
   InitColTitres( fgLiensDirigNew_f, 7, 'TOB;Nom;TOB;Code;Fonction;Infos;Tri' );
   InitColTitres( fgLiensAutreNew_f, 7, 'TOB;Nom;TOB;Code;Fonction;Infos;Tri' );
   InitColCache( fgLiensTitreNew_f, IntToStr( nLNColToL_f ) + ';' + IntToStr( nLNColToF_f ) + ';' + IntToStr( nLNColTri_f ));
   InitColCache( fgLiensDirigNew_f, IntToStr( nLNColToL_f ) + ';' + IntToStr( nLNColToF_f ) + ';' + IntToStr( nLNColTri_f ));
   InitColCache( fgLiensAutreNew_f, IntToStr( nLNColToL_f ) + ';' + IntToStr( nLNColToF_f ) + ';' + IntToStr( nLNColTri_f ));

   nFOColToF_f := 0;
   nFOColFon_f := 1;

   InitColTitres( fgFonctTitre_f, 2, '' );
   InitColTitres( fgFonctDirig_f, 2, '' );
   InitColCache( fgFonctTitre_f, IntToStr( nFOColToF_f ) );
   InitColCache( fgFonctDirig_f, IntToStr( nFOColToF_f ) );

end;

{*****************************************************************
Auteur ....... : BM
Date ......... : 01/08/02
Procédure .... : InitColTitres
Description .. : Initialise le titres des colonnes de la grille
Paramètres ... : La grille
                 Le nombre de colonnes
                 La liste de titres des colonnes
*****************************************************************}

procedure TOF_TRANSFORM.InitColTitres( var fgGrille_p : THGrid; nColNb_p : integer; sTitres_p : string );
var
   tstTitres_l : HTStrings;
   sTmp_l : string;
   nColNb_l : integer;
begin
   // Nombre de colonnes
   fgGrille_p.ColCount := nColNb_p;

   // Titres des colonnes
   if sTitres_p = '' then
   begin
      fgGrille_p.FixedRows := 0;
   end
   else
   begin
      tstTitres_l := HTStringList.Create ;
      sTmp_l := ReadTokenSt(sTitres_p);
      nColNb_l := 0;

      while ( sTmp_l <> '' ) and ( nColNb_l < nColNb_p ) do
        begin
         tstTitres_l.Add( sTmp_l );
         nColNb_l := nColNb_l + 1;
           sTmp_l := ReadTokenSt(sTitres_p);
        end;

      fgGrille_p.FixedRows := 1;
      fgGrille_p.Titres := tstTitres_l;
      tstTitres_l.Free;

   end;
end;

{*****************************************************************
Auteur ....... : BM
Date ......... : 01/08/02
Procédure .... : InitColCache
Description .. : Cache des colonnes de la grille
Paramètres ... : La grille
                 La liste des colonnes à cacher
*****************************************************************}

procedure TOF_TRANSFORM.InitColCache( var fgGrille_p : THGrid; sColCache_p : string );
var
   sTmp_l : string;
   nColNo_l : integer;
begin
   sTmp_l := ReadTokenSt(sColCache_p);
   while sTmp_l <> '' do
   begin
      nColNo_l:= StrToInt( sTmp_l );
      if ( nColNo_l >= 0 ) and  ( nColNo_l < fgGrille_p.ColCount )then
      begin
           fgGrille_p.ColWidths[ nColNo_l ] := -1;
      end;
      sTmp_l := ReadTokenSt(sColCache_p);
   end;
end;

{*****************************************************************
Auteur ....... : BM
Date ......... : 25/07/02
Procédure .... : AfficherLiens
Description .. : Affiche la liste des liens dans la grille
Paramètres ... : La TOB de la liste
                 La grille
*****************************************************************}

procedure TOF_TRANSFORM.AfficherLiens( TOBLiens_p : TOB; var fgLiens_p : THGrid );
var
   nLiensInd_l : integer;
   sInfo_l : string;
   tnColLargeur_l, tnNewColLargeur_l : array of integer;
begin
   SetLength( tnColLargeur_l, fgLiens_p.ColCount );
   if TOBLiens_p.Detail.Count > 0 then
   begin
      fgLiens_p.RowCount := TOBLiens_p.Detail.Count +  fgLiens_p.FixedRows;
      ColGetLargeur( fgLiens_p, tnColLargeur_l );

      for nLiensInd_l := 0 to TOBLiens_p.Detail.Count - 1 do
        begin
           fgLiens_p.Cells[nLAColToL_f, nLiensInd_l + fgLiens_p.FixedRows] := IntToStr(nLiensInd_l);
           fgLiens_p.Cells[nLAColNom_f, nLiensInd_l + fgLiens_p.FixedRows] := TOBLiens_p.Detail[nLiensInd_l].GetValue('ANL_NOMPER' );
           fgLiens_p.Cells[nLAColCod_f, nLiensInd_l + fgLiens_p.FixedRows] := TOBLiens_p.Detail[nLiensInd_l].GetValue('ANL_FONCTION' );
           fgLiens_p.Cells[nLAColFon_f, nLiensInd_l + fgLiens_p.FixedRows] := TOBLiens_p.Detail[nLiensInd_l].GetValue('ANL_AFFICHE' );
         if TOBLiens_p.Detail[nLiensInd_l].GetValue('ANL_INFO' ) = null then
            sInfo_l := ''
         else
            sInfo_l := TOBLiens_p.Detail[nLiensInd_l].GetValue('ANL_INFO' );
           fgLiens_p.Cells[nLAColInf_f, nLiensInd_l + fgLiens_p.FixedRows] := sInfo_l;
           fgLiens_p.Cells[nLAColTri_f, nLiensInd_l + fgLiens_p.FixedRows] := TOBLiens_p.Detail[nLiensInd_l].GetValue('ANL_TRI' );
           TOBLiens_p.Detail[nLiensInd_l].AddChampSup( 'NEWFONCTION', True);
           TOBLiens_p.Detail[nLiensInd_l].AddChampSup( 'OPERATION', True);
         TOBLiens_p.Detail[nLiensInd_l].PutValue('OPERATION', 'S' );
         ColNewLargeur( fgLiens_p, nLiensInd_l + fgLiens_p.FixedRows, tnColLargeur_l );
      end;

      ColSetLargeur( fgLiens_p, tnColLargeur_l );
      SetLength( tnNewColLargeur_l, THGrid( GetControl( fgLiens_p.Name + 'NEW' )).ColCount );

      tnNewColLargeur_l[nLNColToL_f] := tnColLargeur_l[nLAColToL_f];
      tnNewColLargeur_l[nLNColNom_f] := tnColLargeur_l[nLAColNom_f];
      tnNewColLargeur_l[nLNColToF_f] := tnColLargeur_l[nLAColToL_f];;
      tnNewColLargeur_l[nLNColCod_f] := tnColLargeur_l[nLAColCod_f];
      tnNewColLargeur_l[nLNColFon_f] := tnColLargeur_l[nLAColNom_f];
      tnNewColLargeur_l[nLNColInf_f] := tnColLargeur_l[nLAColInf_f];
      tnNewColLargeur_l[nLNColTri_f] := tnColLargeur_l[nLAColInf_f];

      ColSetLargeur( THGrid( GetControl( fgLiens_p.Name + 'NEW' )), tnNewColLargeur_l );
   end;

end;

{*****************************************************************
Auteur ....... : BM
Date ......... : 25/07/02
Procédure .... : AfficherFonctions
Description .. : Affiche la liste des fonctions dans la grille
Paramètres ... : La TOB de la liste
                 La grille
*****************************************************************}

procedure TOF_TRANSFORM.AfficherFonctions( TOBLiens_p : TOB; var fgFonction_p : THGrid );
var
   nTOBFonctInd_l : integer;
   tnColLargeur_l : array of integer;
begin
   SetLength( tnColLargeur_l, fgFonction_p.ColCount );
   fgFonction_p.RowCount := TOBLiens_p.Detail.Count + fgFonction_p.FixedRows;
   if TOBLiens_p.Detail.Count > 0 then
   begin
      ColGetLargeur( fgFonction_p, tnColLargeur_l );
      for nTOBFonctInd_l := 0 to TOBLiens_p.Detail.Count - 1 do
        begin
           fgFonction_p.Cells[nFOColToF_f, nTOBFonctInd_l + fgFonction_p.FixedRows] := IntToStr(nTOBFonctInd_l);
           fgFonction_p.Cells[nFOColFon_f, nTOBFonctInd_l + fgFonction_p.FixedRows] := TOBLiens_p.Detail[nTOBFonctInd_l].GetValue('JTF_FONCTABREGE' );
         ColNewLargeur( fgFonction_p, nTOBFonctInd_l + fgFonction_p.FixedRows, tnColLargeur_l );
      end;
      ColSetLargeur( fgFonction_p, tnColLargeur_l );
   end;
end;

{*****************************************************************
Auteur ....... : BM
Date ......... : 26/07/02
Procédure .... : ColGetLargeur
Description .. : Largeur initiale des colonnes
Paramètres ... : La grille
                 Le tableau de largeurs
*****************************************************************}

procedure TOF_TRANSFORM.ColGetLargeur( fgGrille_p : THGrid; var tnColLargeur_p : array of integer );
var
   nColInd_l : integer;
begin
   for nColInd_l := 0 to fgGrille_p.ColCount - 1 do
   begin
      tnColLargeur_p[nColInd_l] := fgGrille_p.ColWidths[ nColInd_l];
     end;
end;

{*****************************************************************
Auteur ....... : BM
Date ......... : 26/07/02
Procédure .... : ColNewLargeur
Description .. : Calcule la largeur des colonnes en fonction du contenu
Paramètres ... : La grille
                 L'indice de la ligne courante
                 Le tableau de largeurs
*****************************************************************}

procedure TOF_TRANSFORM.ColNewLargeur( fgGrille_p : THGrid; nIndice_p : integer; var tnColLargeur_p : array of integer );
var
   nColInd_l, nColLargeur_l : integer;

begin
   for nColInd_l := 0 to fgGrille_p.ColCount - 1 do
   begin
      if tnColLargeur_p[nColInd_l] <> -1 then
      begin
         // Colonne non cachées
           nColLargeur_l := fgGrille_p.Canvas.TextWidth(fgGrille_p.Cells[ nColInd_l, nIndice_p ]) + 8;
           if nColLargeur_l > tnColLargeur_p[nColInd_l] then
              tnColLargeur_p[nColInd_l] := nColLargeur_l;
      end;
     end;
end;

{*****************************************************************
Auteur ....... : BM
Date ......... : 26/07/02
Procédure .... : ColSetLargeur
Description .. : Fixe la largeur des colonnes
Paramètres ... : La grille
                 Le tableau de largeurs
*****************************************************************}

procedure TOF_TRANSFORM.ColSetLargeur( fgGrille_p : THGrid; var tnColLargeur_p : array of integer );
var
   nColInd_l : integer;

begin
   for nColInd_l := 0 to fgGrille_p.ColCount - 1 do
   begin
      fgGrille_p.ColWidths[nColInd_l] := tnColLargeur_p[nColInd_l];
   end;
end;

{*****************************************************************
Auteur ....... : BM
Date ......... : 26/07/02
Procédure .... : ValiderBoutons
Description .. : Valide ou dévalide les boutons SUPPRIMER CHANGER ANNULER
Paramètres ... : Le nom de la grille des liens
                 La ligne du lien sélectionné
                 Le nombre de lignes de la grille des liens
                 L'indice de la tob associée
                 La ligne de la fonction sélectionnée
                 Le nombre de lignes de la grille des fonctions
                 L'indice de la tob associée
*****************************************************************}

procedure TOF_TRANSFORM.ValiderBoutons( sNom_p : string; nLiensRow_p, nLiensFixedRow_p, nLiensRowCount_p : integer; sLiensVal_p : string; nFonctRow_p, nFonctFixedRow_p, nFonctRowCount_p : integer; sFonctVal_p : string );

begin
   TToolBarButton97( GetControl( 'B_SUP' + sNom_p )).Enabled    := ValiderLeBouton( 'B_SUP'    + sNom_p, true, nLiensRow_p, nLiensFixedRow_p, nLiensRowCount_p, sLiensVal_p );
   TToolBarButton97( GetControl( 'B_CHANGE' + sNom_p )).Enabled := ValiderLeBouton( 'B_CHANGE' + sNom_p, TToolBarButton97( GetControl( 'B_SUP' + sNom_p )).Enabled, nFonctRow_p, nFonctFixedRow_p, nFonctRowCount_p, sFonctVal_p );
end;


function TOF_TRANSFORM.ValiderLeBouton( sNomBouton_p : string; bForce_p : boolean; nRow_p, nRowFixed, nRowCount_p : integer; sTobVal_p : string ) : boolean;
begin
   if ( nRowFixed = -1 ) and ( nRowCount_p = - 1 ) and bForce_p then
      TToolBarButton97( GetControl( sNomBouton_p )).Enabled := true
   else
   if ( nRow_p >= nRowFixed ) and ( nRow_p < nRowCount_p )
      and bForce_p and ( sTobVal_p <> '' ) then
      TToolBarButton97( GetControl( sNomBouton_p )).Enabled := true
   else
      TToolBarButton97( GetControl( sNomBouton_p )).Enabled := False;
   result := TToolBarButton97( GetControl( sNomBouton_p )).Enabled;
end;

{*****************************************************************
Auteur ....... : BM
Date ......... : 26/07/02
Procédure .... : Supprimer
Description .. : Supprime un liens de la grille de gauche
Paramètres ... : La grille de liens actuels
                 La TOB des opérations effectuées
*****************************************************************}

procedure TOF_TRANSFORM.Supprimer( var fgLiens_p : THGrid );
var
   sNomBouton_l : string;
begin
   sNomBouton_l := Copy( fgLiens_p.Name, Length(fgLiens_p.Name) - 4, Length(fgLiens_p.Name) - 1 );
   EnleveLigne( fgLiens_p, fgLiens_p.Row );
   LigneOnSelect( fgLiens_p);
end;

{*****************************************************************
Auteur ....... : BM
Date ......... : 26/07/02
Procédure .... : Changer
Description .. : Change la fonction d'un lien
Paramètres ... : La grille des liens actuels et les passe dans la grille de droite
                 La grille des fonctions
                 La grille des nouveaux liens
                 La TOB des opéarations effectuées
*****************************************************************}

procedure TOF_TRANSFORM.Changer( var fgLiens_p : THGrid; fgFonction_p : THGrid; var fgLiensNew_p : THGrid; TOBFonctions_p : TOB );
var
   sCodeFonction_l, sNomFonction_l, sLNColToF_l : string;
   nLienNewInd_l, nFonctNewInd_l : integer;
begin
   if fgLiens_p.Name = 'LIENSAUTRE' then
   begin
       sLNColToF_l := '-1';
       sCodeFonction_l := fgLiens_p.Cells[nLAColCod_f, fgLiens_p.Row];
       sNomFonction_l := fgLiens_p.Cells[nLAColFon_f, fgLiens_p.Row];
   end
   else
   begin
       nFonctNewInd_l := StrToInt( fgFonction_p.Cells[nFOColToF_f, fgFonction_p.Row] );
       sLNColToF_l := IntToStr( nFonctNewInd_l);
       sCodeFonction_l := TOBFonctions_p.Detail[nFonctNewInd_l].GetValue('JFT_FONCTION');
       sNomFonction_l := TOBFonctions_p.Detail[nFonctNewInd_l].GetValue('JTF_AFFICHE');
   end;

   if VerifLien( fgLiens_p.Cells[nLAColNom_f, fgLiens_p.Row], sCodeFonction_l, fgLiensNew_p ) then
   begin
      nLienNewInd_l := AjouteLigne( fgLiensNew_p, fgLiensNew_p.RowCount );

      fgLiensNew_p.Cells[nLNColToL_f, nLienNewInd_l] := fgLiens_p.Cells[nLAColToL_f, fgLiens_p.Row];
      fgLiensNew_p.Cells[nLNColNom_f, nLienNewInd_l] := fgLiens_p.Cells[nLAColNom_f, fgLiens_p.Row];
      fgLiensNew_p.Cells[nLNColToF_f, nLienNewInd_l] := sLNColToF_l;
      fgLiensNew_p.Cells[nLNColCod_f, nLienNewInd_l] := sCodeFonction_l;
      fgLiensNew_p.Cells[nLNColFon_f, nLienNewInd_l] := sNomFonction_l;
      fgLiensNew_p.Cells[nLNColInf_f, nLienNewInd_l] := fgLiens_p.Cells[nLAColInf_f, fgLiens_p.Row];
      fgLiensNew_p.Cells[nLNColTri_f, nLienNewInd_l] := fgLiens_p.Cells[nLAColTri_f, fgLiens_p.Row];
      fgLiensNew_p.SortGrid(nLNColTri_f, false);

      EnleveLigne( fgLiens_p, fgLiens_p.Row );

      LigneOnSelect( fgLiens_p);
      LigneNewOnSelect( fgLiensNew_p);
   end
   else
   begin
      PGIInfo(
              'L''individu ' + fgLiens_p.Cells[nLAColNom_f, fgLiens_p.Row] +
              ' existe déjà avec la fonction ' + sNomFonction_l +
               '.#13#10' +
               'Vous devez le supprimer ou lui affecter une nouvelle fonction.#13#10',
               Ecran.Caption );

   end;
end;

{*****************************************************************
Auteur ....... : BM
Date ......... : 26/07/02
Procédure .... : Annuler
Description .. : Annule la dernière opération
Paramètres ... : La TOB des opérations effectuées
                 La TOB des liens d'origine
                 La grille des liens actuels
                 La grille des nouveaux liens

*****************************************************************}

procedure TOF_TRANSFORM.Annuler( TOBLiens_p : TOB; var fgLiens_p : THGrid; var fgLiensNew_p : THGrid );
var
   nLienInd_l, nTOBLienInd_l : integer;
begin
   nTOBLienInd_l := StrToInt( fgLiensNew_p.Cells[nLNColToL_f, fgLiensNew_p.Row] );
   EnleveLigne( fgLiensNew_p, fgLiensNew_p.Row );
   nLienInd_l := AjouteLigne( fgLiens_p, fgLiens_p.RowCount );

   fgLiens_p.Cells[nLAColToL_f, nLienInd_l] := IntToStr( nTOBLienInd_l );
   fgLiens_p.Cells[nLAColNom_f, nLienInd_l] := TOBLiens_p.Detail[nTOBLienInd_l].GetValue('ANL_NOMPER' );
   fgLiens_p.Cells[nLAColCod_f, nLienInd_l] := TOBLiens_p.Detail[nTOBLienInd_l].GetValue('ANL_FONCTION' );
   fgLiens_p.Cells[nLAColFon_f, nLienInd_l] := TOBLiens_p.Detail[nTOBLienInd_l].GetValue('ANL_AFFICHE' );
   fgLiens_p.Cells[nLAColInf_f, nLienInd_l] := TOBLiens_p.Detail[nTOBLienInd_l].GetValue('ANL_INFO' );
   fgLiens_p.Cells[nLAColTri_f, nLienInd_l] := TOBLiens_p.Detail[nTOBLienInd_l].GetValue('ANL_TRI' );
   fgLiens_p.SortGrid(nLAColTri_f, false);

   LigneOnSelect( fgLiens_p );
   LigneNewOnSelect( fgLiensNew_p );

end;
{*****************************************************************
Auteur ....... : BM
Date ......... : 29/07/02
Procédure .... : EnleveLigne
Description .. : Enlève une ligne dans une grille
Paramètres ... : La grille
                 L'indice de la ligne
*****************************************************************}

procedure TOF_TRANSFORM.EnleveLigne( var fgGrille_p : THGrid; nRow_p : integer );
var nColInd_l : integer;
begin
   if ( fgGrille_p.RowCount - 1 = fgGrille_p.FixedRows ) then
   begin
      for nColInd_l := 0 to fgGrille_p.ColCount - 1 do
      begin
         fgGrille_p.Cells[nColInd_l, nRow_p] := '';
      end;
   end
   else
   begin
      fgGrille_p.DeleteRow( nRow_p );
   end;
end;

{*****************************************************************
Auteur ....... : BM
Date ......... : 29/07/02
Fonction ..... : AjouteLigne
Description .. : Ajoute une ligne dans une grille
Paramètres ... : La grille
                 L'indice de la ligne
Renvoie ...... : L'indice de la ligne ajoutée
*****************************************************************}

function TOF_TRANSFORM.AjouteLigne( var fgGrille_p : THGrid; nRow_p : integer ) : integer;
begin
   if ( fgGrille_p.RowCount - 1 = fgGrille_p.FixedRows ) and ( fgGrille_p.Cells[0,nRow_p - 1] = '' ) then
      nRow_p := nRow_p - 1
   else
      fgGrille_p.InsertRow( nRow_p );
   result := nRow_p;
end;
{*****************************************************************
Auteur ....... : BM
Date ......... : 19/09/02
Fonction ..... : VerifLien
Description .. : Vérifie si le lien n'existe pas déjà dans le tableau
Paramètres ... : Le nom de la personne
                 Le code de la fonction
                 La grille des nouveaux liens
Renvoie ...... : True si la ligne n'existe pas
                 False sinon                 
*****************************************************************}

function TOF_TRANSFORM.VerifLien( sNomPer_p, sCodeFonction_p : string; fgLiensNew_p : THGrid ) : boolean;
var
   nLiensInd_l : integer;
   bTrouve_l : boolean;
begin
   bTrouve_l := true;
   nLiensInd_l := fgLiensNew_p.FixedRows;
   while ( bTrouve_l ) and ( nLiensInd_l < fgLiensNew_p.RowCount ) do
   begin
      if ( fgLiensNew_p.Cells[nLNColNom_f, nLiensInd_l] = sNomPer_p )
      and ( fgLiensNew_p.Cells[nLNColCod_f, nLiensInd_l] = sCodeFonction_p ) then
         bTrouve_l := false;
      nLiensInd_l := nLiensInd_l + 1;
   end;
   result := bTrouve_l;
end;


{*****************************************************************
Auteur ....... : BM
Date ......... : 30/07/02
Fonction ..... : Enregistrable
Description .. : Vérifie si les grilles des liens sont vides et autorise alors l'enregistrement des modifications
Paramètres ... :
Renvoie ...... : True si les grilles sont vides, false sinon
*****************************************************************}

function TOF_TRANSFORM.Enregistrable : boolean;
begin
   if  ( fgLiensTitre_f.FixedRows + 1 = fgLiensTitre_f. RowCount )
   and ( fgLiensDirig_f.FixedRows + 1 = fgLiensDirig_f.RowCount )
   and ( fgLiensAutre_f.FixedRows + 1 = fgLiensAutre_f.RowCount ) then
      result := true
   else
      result := false;
end;

{*****************************************************************
Auteur ....... : BM
Date ......... : 30/07/02
Procedure .... : Enregistrer
Description .. : Enregistre les modifications dans la table ANNULIEN
Paramètres ... :
*****************************************************************}

procedure TOF_TRANSFORM.Enregistrer;
begin
   EcranVersClasse( fgLiensTitreNew_f, TOBLiensTitre_f );
   EcranVersClasse( fgLiensDirigNew_f, TOBLiensDirig_f );
   EcranVersClasse( fgLiensAutreNew_f, TOBLiensAutre_f );
   MiseAJour( TOBLiensTitre_f, TOBFonctTitre_f );
   MiseAJour( TOBLiensDirig_f, TOBFonctDirig_f );
   MiseAJour( TOBLiensAutre_f, nil );
end;
{*****************************************************************
Auteur ....... : BM
Date ......... : 30/07/02
Procédure .... : EcranVersClasse
Description .. : Met la TOB à jour avec les informations de la grille
Paramètres ... : La grille des nouveaux liens
                 La TOB de liens

*****************************************************************}

procedure TOF_TRANSFORM.EcranVersClasse( fgLiensNew_p : THGrid; var TOBLiens_p : TOB );
var
   nLiensInd_l, nTOBLienInd_l, nTOBFoncInd_l : integer;
begin
   for nLiensInd_l := fgLiensNew_p.FixedRows to fgLiensNew_p.RowCount - 1 do
   begin
      if fgLiensNew_p.Cells[nLNColToL_f, nLiensInd_l] <> '' then
      begin
         nTOBLienInd_l := StrToInt( fgLiensNew_p.Cells[nLNColToL_f, nLiensInd_l] );
         //
         nTOBFoncInd_l := StrToInt( fgLiensNew_p.Cells[nLNColToF_f, nLiensInd_l] );
           TOBLiens_p.Detail[nTOBLienInd_l].PutValue('NEWFONCTION', nTOBFoncInd_l );
         TOBLiens_p.Detail[nTOBLienInd_l].PutValue('OPERATION', 'C' );
      end;
   end;
end;

{*****************************************************************
Auteur ....... : BM
Date ......... : 30/07/02
Fonction ..... : MiseAJour
Description .. : Met la base à jour
Paramètres ... : La TOB
Renvoie ...... : True si la maj s'est bien passée
*****************************************************************}

procedure TOF_TRANSFORM.MiseAJour( TOBLiens_p, TOBFonctions_p : TOB );
var
   nLiensInd_l, nTOBFoncInd_l : integer;
   sRequete_l, sNouFormeCode_l : string;
   TOB_l : TOB;
begin

   for nLiensInd_l := 0 to TOBLiens_p.Detail.Count - 1 do
   begin
      // Suppression
      if TOBLiens_p.Detail[nLiensInd_l].GetValue('OPERATION' ) = 'S' then
      begin

         sRequete_l :=
               'delete ' +
               'from   ANNULIEN ' +
               'where  ANL_GUIDPERDOS = "' + sGuidPerDos_fp + '" ' +
               '  and  ANL_GUIDPER    = "' + TOBLiens_p.Detail[nLiensInd_l].GetValue('ANL_GUIDPER') + '" ' +
               '  and  ANL_TYPEDOS    = "STE" ' +
               '  and  ANL_FONCTION   = "' + TOBLiens_p.Detail[nLiensInd_l].GetValue('ANL_FONCTION' ) + '" ';

           ExecuteSQL ( sRequete_l );

      end;

      // Changement ou conservation
      if ( TOBLiens_p.Detail[nLiensInd_l].GetValue('OPERATION' ) = 'C' ) then
      begin
         nTOBFoncInd_l := TOBLiens_p.Detail[nLiensInd_l].GetValue('NEWFONCTION' );

         if nTOBFoncInd_l = -1 then
         begin
            // Bouton CONSERVER, pas de nouvelle fonction, on conserve l'ancienne
            TOB_l := TOB.Create( '', nil, -1 );
            sRequete_l :=
                     'select JFT_DEFAUT ' +
                     'from   JUFONCTION ' +
                     'where  JFT_FORME = "' + sAncFormeCode_fp + '" ' +
                     '  and  JFT_FONCTION   = "' + TOBLiens_p.Detail[nLiensInd_l].GetValue('ANL_FONCTION' ) + '" ';
            TOB_l.LoadDetailFromSQL(sRequete_l);

            sNouFormeCode_l := sNouFormeCode_fp;
            if TOB_l.Detail.Count>0 then
            begin
                if TOB_l.Detail[0].GetValue('JFT_DEFAUT' ) <> 'X' then
                   sNouFormeCode_l := '';
            end;

            TOB_l.Free;

            sRequete_l := 'update ANNULIEN ' +
              'set    ANL_APPMODIF   = "TRANSFO",' +
              '       ANL_FORME      = "' + sNouFormeCode_l + '" ' +
              'where  ANL_GUIDPERDOS = "' + sGuidPerDos_fp + '" ' +
              '  and  ANL_GUIDPER    = "' + TOBLiens_p.Detail[nLiensInd_l].GetValue('ANL_GUIDPER') + '" ' +
              '  and  ANL_TYPEDOS    = "STE" ' +
              '  and  ANL_FONCTION   = "' + TOBLiens_p.Detail[nLiensInd_l].GetValue('ANL_FONCTION' ) + '" ';

              ExecuteSQL ( sRequete_l );

         end
         else
         begin
            // Bouton CHANGER
            if TOBFonctions_p.Detail[nTOBFoncInd_l].GetValue('JFT_DEFAUT' ) <> 'X' then
               sNouFormeCode_l := ''
            else
               sNouFormeCode_l := sNouFormeCode_fp;

            sRequete_l := 'update ANNULIEN ' +
              'set    ANL_APPMODIF   = "TRANSFO",' +
              '       ANL_FONCTION   = "' + TOBFonctions_p.Detail[nTOBFoncInd_l].GetValue('JFT_FONCTION' ) + '",' +
              '       ANL_FORME      = "' + sNouFormeCode_l + '", ' +
              '       ANL_RACINE     = "' + TOBFonctions_p.Detail[nTOBFoncInd_l].GetValue('JFT_RACINE' ) + '",' +
              '       ANL_AFFICHE    = "' + TOBFonctions_p.Detail[nTOBFoncInd_l].GetValue('JTF_AFFICHE' ) + '", ' +
              '       ANL_TRI        = "' + IntToStr( TOBFonctions_p.Detail[nTOBFoncInd_l].GetValue('JFT_TRI' ) ) + '" ' +
              'where  ANL_GUIDPERDOS = "' + sGuidPerDos_fp + '" ' +
              '  and  ANL_GUIDPER    = "' + TOBLiens_p.Detail[nLiensInd_l].GetValue('ANL_GUIDPER') + '" ' +
              '  and  ANL_TYPEDOS    = "STE" ' +
              '  and  ANL_FONCTION   = "' + TOBLiens_p.Detail[nLiensInd_l].GetValue('ANL_FONCTION' ) + '" ';

              ExecuteSQL ( sRequete_l );

         end;
      end;
     end;
end;

{*****************************************************************
* Initialisation                                                  *
*****************************************************************}
Initialization
  registerclasses ( [ TOF_TRANSFORM ] ) ;
end.

