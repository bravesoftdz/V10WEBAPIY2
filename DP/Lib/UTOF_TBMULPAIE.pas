Unit UTOF_TBMULPAIE;

Interface

Uses StdCtrls,
     Controls,
     Classes,
     windows,
{$IFDEF EAGLCLIENT}
     eMul,
     uTob,
     MaineAGL,
{$ELSE}
     db,
     {$IFNDEF DBXPRESS}dbtables,{$ELSE}uDbxDataSet,{$ENDIF}
     mul,
     Fe_Main,
     HDB,
{$ENDIF}
     Hqry,
     forms,
     sysutils,
     ComCtrls,
     HCtrls,
     HEnt1,
     HMsgBox,
     UTOF,
     GalOutil,
     HTB97,
     UTOF_TBSTATPAIE;

const
  TabMois : array [1..12] of string = ('Janvier','Fevrier','Mars','Avril','Mai','Juin','Juillet','Aout','Septembre','Octobre','Novembre','Decembre');

Type
  TOF_TBMULPAIE = Class (TOF)
    procedure OnArgument (S : String ) ; override ;
    procedure Aff_TBFicheDossier (Sender : TObject);
    procedure InitialiserComboBoxMois;
    procedure InitialiserComboBoxAnnee;
    procedure Supprimer (Sender : TObject);
    procedure Imprimer (Sender : TObject);
    procedure FormKeyDown (Sender : TObject; var key:Word; Shift:TShiftState);
    procedure BCHERCHE_OnClick(Sender: TObject);
  end ;

//---------------------------------
//--- Déclaration des procédures
//---------------------------------
procedure Aff_TBMulPaie ;

Implementation

//------------------------------------------------------
//--- Nom   : Aff_DPMulPaie
//--- Objet : Affichage de la fiche Multicritére Paie
//------------------------------------------------------
procedure Aff_TBMulPaie ;
begin
 AGLLanceFiche('DP','TBMULPAIE','','','');
end;

//---------------------------------------------------
//--- Nom   : OnArgument
//--- Objet : Initialisation de la fiche TBMULPAIE
//---------------------------------------------------
procedure TOF_TBMULPAIE.OnArgument (S : String );
begin
 Inherited ;
 TFMul (Ecran).Caption:='Tableau de bord paie';
 UpdateCaption (Ecran);

 //--- Gestion des évenements
 THEdit (GetControl ('DT1_NODOSSIER')).OnElipsisClick:=Aff_TBFicheDossier;
 TToolbarButton97(GetControl('BSUPPRIMER')).Onclick:=Supprimer;
 TToolbarButton97(GetControl('BIMPRIMER')).Onclick:=Imprimer;
 TToolBarButton97(GetControl('BCHERCHE')).OnClick := BCHERCHE_OnClick;
 Ecran.OnkeyDown:=FormKeyDown;

 InitialiserComboGroupeConf (THMultiValCombobox(GetControl('GROUPECONF')));

 //--- Initialisation des ComboBox
 InitialiserComboBoxMois;
 InitialiserComboBoxAnnee;

 //--- Oubli dans la fiche
 SetControlVisible('bSelectAll', True);
end ;

//---------------------------------------
//--- Nom   : Aff_TBFicheDossier
//--- Objet : Affiche la fiche dossier
//---------------------------------------
procedure TOF_TBMULPAIE.Aff_TBFicheDossier (Sender :TObject);
var Chaine  : String;
    SauverDossier : String;
begin
 SauverDossier:=GetControlText('DT1_NODOSSIER');
 Chaine:=AGLLanceFiche('YY','YYDOSSIER_SEL', '','',GetControlText('DT1_NODOSSIER'));
 if Chaine<>'' then
  SetControlText ('DT1_NODOSSIER', READTOKENST(Chaine))
 else
  SetControlText('DT1_NODOSSIER', SauverDossier);
end;

//-----------------------------
//--- Nom : BCherche_OnClick
//-----------------------------
procedure TOF_TBMULPAIE.BCHERCHE_OnClick(Sender: TObject);
var ChXXWhere : String;
begin
 ChXXWhere:=GererCritereGroupeConf (THMultiValCombobox(GetControl('GROUPECONF')),TcheckBox (GetControl ('SANSGRPCONF')).Checked);
 ChXXWhere:=ChXXWhere+GererCritereDivers ();
 // $$$ JP 04/09/06 - plus autorisé en D7/Unicode ... TEdit(GetControl('XX_WHERE')).Text:=ChXXWhere;
 SetControlText ('XX_WHERE', chXXWhere);

 // Traitement générique
 TFMul(Ecran).BChercheClick(Sender);
end;

//----------------------------------------------------------
//--- Nom   : InitialiserComboBoxMois
//--- Objet : Initialisation des éléments de la combo box
//----------------------------------------------------------
procedure TOF_TBMULPAIE.InitialiserComboBoxMois;
var ChSql  : String;
    RSql   : Tquery;
    Indice : Integer;
    Nombre : Integer;
begin
 ChSql:='Select DT1_MOIS from DPTABGENPAIE where group by DT1_MOIS';
  try
   RSql:=Opensql (ChSql, TRUE);
   RSql.First;
   Nombre:=RSql.recordCount;

   //THValComboBox (GetControl ('DT1_MOIS')).Items.AddObject ('<<Tous>>',nil);
   //THValComboBox (GetControl ('DT1_MOIS')).Values.Add ('');

   for Indice:=0 to Nombre-1 do
    begin
     THMultiValComboBox (GetControl ('DT1_MOIS')).Items.AddObject (TabMois [StrToInt (RSql.FindField ('DT1_MOIS').AsString)],nil);
     THMultiValComboBox (GetControl ('DT1_MOIS')).Values.Add (RSql.FindField ('DT1_MOIS').AsString);
     RSql.Next;
    end;

   //if (Nombre>0) then
   // THMultiValComboBox (GetControl ('DT1_MOIS')).Text:=THMultiValComboBox (GetControl ('DT1_MOIS')).Items [0];

   Ferme (RSql);
  except
   begin
    PGIINFO ('Problème dans la table DPTABGENPAIE','Erreur');
   end;
  end;
end;

//----------------------------------------------------------
//--- Nom   : InitialiserComboBoxAnnee
//--- Objet : Initialisation des éléments de la combo box
//----------------------------------------------------------
procedure TOF_TBMULPAIE.InitialiserComboBoxAnnee;
var ChSql   : String;
    RSql   : Tquery;
    Indice : Integer;
    Nombre : Integer;
begin
 ChSql:='Select DT1_ANNEE from DPTABGENPAIE where group by DT1_ANNEE';
  try
   RSql:=Opensql (ChSql, TRUE);
   RSql.First;
   Nombre:=RSql.recordCount;

   for Indice:=0 to Nombre-1 do
    begin
     THMultiValComboBox (GetControl ('DT1_ANNEE')).Items.AddObject (RSql.FindField ('DT1_ANNEE').AsString,nil);
     THMultiValComboBox (GetControl ('DT1_ANNEE')).Values.Add (RSql.FindField ('DT1_ANNEE').AsString);
     RSql.Next;
    end;

   //if (Nombre>0) then
   // THMultiValComboBox (GetControl ('DT1_ANNEE')).Text:=THMultiValComboBox (GetControl ('DT1_ANNEE')).Items [0];

   Ferme (RSql);
  except
   begin
    PGIINFO ('Problème dans la table DPTABGENPAIE','Erreur');
   end;
  end;
end;

//-------------------------------------------
//--- Nom   : Supprimer
//--- Objet : Suppression des informations
//-------------------------------------------
procedure TOF_TBMULPAIE.Supprimer (Sender : TObject);
var
{$IFDEF EAGLCLIENT}
    Liste    : THGrid;
{$ELSE}
    Liste    : THDBGrid;
{$ENDIF}
    RSql     : THQuery;
    ChChaine : String;
    Indice   : Integer;
begin
{$IFDEF EAGLCLIENT}
 Liste:=THGrid (GetControl ('FListe'));
{$ELSE}
 Liste:=THDBGrid (GetControl ('FListe'));
{$ENDIF}

 if (Liste.NbSelected=0) and (not liste.AllSelected) then
  begin
   PGIINFO('Aucun élément sélectionné','Information');
   exit;
  end
 else
  begin
   ChChaine:='Vous allez supprimer définitivement les informations.#13#10Confirmez vous l''opération ?';
   if HShowMessage('0;Suppression;'+ChChaine+';Q;YN;N;N;','','')<>mrYes then
    exit ;

   RSql:=TFMul(Ecran).Q;
   if (Liste.AllSelected) then
    begin
{$IFDEF EAGLCLIENT}
     if not TFMul(Ecran).Fetchlestous then
      PGIInfo('Impossible de récupérer tous les enregistrements')
     else
{$ENDIF}
      begin
       RSql.First;
       while not RSql.Eof do
        begin
         ExecuteSQL('DELETE FROM DPTABCOMPTA WHERE DTC_NODOSSIER="'+RSql.FindField ('DTC_NODOSSIER').AsString+'" And DTC_LIBEXERCICE="'+RSql.FindField ('DTC_LIBEXERCICE').AsString+'"');
         RSql.Next;
        end;
      end;
    end
   else
    begin
     for Indice:=0 to Liste.NbSelected-1 do
      begin
       Liste.GotoLeBookmark(Indice);
{$IFDEF EAGLCLIENT}
       RSql.TQ.Seek(Liste.Row - 1) ;
{$ENDIF}
       ExecuteSQL('DELETE FROM DPTABGENPAIE WHERE DT1_NODOSSIER="'+RSql.FindField ('DT1_NODOSSIER').AsString+'" And DT1_MOIS="'+RSql.FindField ('DT1_MOIS').AsString+'" and DT1_ANNEE="'+RSql.FindField ('DT1_ANNEE').AsString+'"');
      end;
    end;
   TFMul(Ecran).BChercheClick(Nil);
  end;
end;

//-----------------------------------
//--- Nom   : Imprimer
//--- Objet : Imprimer Statistique
//-----------------------------------
procedure TOF_TBMULPAIE.Imprimer (Sender : TObject);
begin
 Aff_TBStatPaie (TFMul(Ecran).Q.SQL.Text);
end;

//-------------------------------------------
//--- Nom   : FormKeyDown
//--- Objet : Scrute l'appuit d'une touche
//-------------------------------------------
procedure TOF_TBMULPAIE.FormKeyDown (Sender : TObject; var key:Word; Shift:TShiftState);
begin
 TFMul(Ecran).FormKeyDown(Sender, Key, Shift);
 case key of
  80,112 : if (Shift=[ssCtrl]) then
            Aff_TBStatPaie (TFMul(Ecran).Q.SQL.Text);

  VK_DELETE : if (Shift=[ssCtrl]) then
               Supprimer (Sender);
 end;
end;

Initialization
  registerclasses ( [ TOF_TBMULPAIE ] ) ;
end.

