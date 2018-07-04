Unit UTOF_TBMULCOMPTA;

Interface

Uses StdCtrls,
     Controls,
     Classes,
     Windows,
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
     HQry,
     forms,
     sysutils,
     ComCtrls,
     HCtrls,
     HEnt1,
     HMsgBox,
     UTOF,
     GalOutil,
     HDB,
     HTB97,
     stat,
     UTOF_TBSTATCOMPTA;

Type
  TOF_TBMULCOMPTA = Class (TOF)
    procedure OnArgument (S : String ) ; override ;
    procedure Aff_TBFicheDossier (Sender : TObject);
    procedure InitialiserComboBoxCompta;
    procedure Imprimer (Sender : TObject);
    procedure FormKeyDown (Sender : TObject; var key:Word; Shift:TShiftState);
    procedure BCHERCHE_OnClick(Sender: TObject);

  end ;

//---------------------------------
//--- Déclaration des procédures
//---------------------------------
procedure Aff_TBMulCompta ;

Implementation

//--------------------------------------------------------
//--- Nom   : Aff_DPMulCompta
//--- Objet : Affichage de la fiche Multicritére Compta
//--------------------------------------------------------
procedure Aff_TBMulCompta ;
begin
 AGLLanceFiche('DP','TBMULCOMPTA','','','');
end;

//---------------------------------------------------
//--- Nom   : OnArgument
//--- Objet : Initialisation de la fiche TBMULPAIE
//---------------------------------------------------
procedure TOF_TBMULCOMPTA.OnArgument (S : String );
begin
 Inherited ;
 TFMul (Ecran).Caption:='Tableau de bord comptable';
 UpdateCaption (Ecran);

 //--- Gestion des évenements
 THEdit (GetControl ('DTC_NODOSSIER')).OnElipsisClick:=Aff_TBFicheDossier;
 SetControlVisible('BSUPPRIMER', False); // FQ 11739
 TToolbarButton97(GetControl('BIMPRIMER')).Onclick:=Imprimer;
 TToolBarButton97(GetControl('BCHERCHE')).OnClick := BCHERCHE_OnClick;
 Ecran.OnkeyDown:=FormKeyDown;

 InitialiserComboGroupeConf (THMultiValCombobox(GetControl('GROUPECONF')));

 //--- Initialisation des ComboBox
 InitialiserComboBoxCompta;

  //--- Oublis dans la fiche
 SetControlVisible('bSelectAll', True);
end ;

//---------------------------------------
//--- Nom   : Aff_TBFicheDossier
//--- Objet : Affiche la fiche dossier
//---------------------------------------
procedure TOF_TBMULCOMPTA.Aff_TBFicheDossier (Sender :TObject);
var Chaine  : String;
    SauverDossier : String;
begin
 SauverDossier:=GetControlText('DTC_NODOSSIER');
 Chaine:=AGLLanceFiche('YY','YYDOSSIER_SEL', '','',GetControlText('DTC_NODOSSIER'));
 if Chaine<>'' then
  SetControlText ('DTC_NODOSSIER', READTOKENST(Chaine))
 else
  SetControlText('DTC_NODOSSIER', SauverDossier);
end;

//-----------------------------
//--- Nom : BCherche_OnClick
//-----------------------------
procedure TOF_TBMULCOMPTA.BCHERCHE_OnClick(Sender: TObject);
var ChXXWhere : String;
begin
 ChXXWhere:=GererCritereGroupeConf (THMultiValCombobox(GetControl('GROUPECONF')),TcheckBox (GetControl ('SANSGRPCONF')).Checked);
 ChXXWhere:=ChXXWhere+GererCritereDivers ();
 // $$$ JP 04/09/06 - plus autorisé en D7/Unicode ... TEdit(GetControl('XX_WHERE')).Text:=ChXXWhere;
 SetControlText ('XX_WHERE', chXXWHere);

 // Traitement générique
 TFMul(Ecran).BChercheClick(Sender);
end;

//----------------------------------------------------------
//--- Nom   : InitialiserComboBoxExercice
//--- Objet : Initialisation des éléments de la combo box
//----------------------------------------------------------
procedure TOF_TBMULCOMPTA.InitialiserComboBoxCompta;
var ChSql   : String;                               
    RSql   : Tquery;
    Indice : Integer;
    Nombre : Integer;
begin
 ChSql:='Select DTC_LIBEXERCICE from DPTABCOMPTA where group by DTC_LIBEXERCICE';
  try
   RSql:=Opensql (ChSql, TRUE);
   RSql.First;
   Nombre:=RSql.recordCount;

   for Indice:=0 to Nombre-1 do
    begin
     THMultiValComboBox (GetControl ('DTC_LIBEXERCICE')).Items.AddObject (RSql.FindField ('DTC_LIBEXERCICE').AsString,nil);
     THMultiValComboBox (GetControl ('DTC_LIBEXERCICE')).Values.Add (RSql.FindField ('DTC_LIBEXERCICE').AsString);
     RSql.Next;
    end;

   //if (Nombre>0) then
   // THMultiValComboBox (GetControl ('DTC_LIBEXERCICE')).Text:=THMultiValComboBox (GetControl ('DTC_LIBEXERCICE')).Items [0];

   Ferme (RSql);
  except
   begin
    PGIINFO ('Problème dans la table DPTABCOMPTA','Erreur');
   end;
  end;
end;


//-----------------------------------
//--- Nom   : Imprimer
//--- Objet : Imprimer Statistique
//-----------------------------------
procedure TOF_TBMULCOMPTA.Imprimer (Sender : TObject);
begin
 Aff_TBStatCompta (TFMul(Ecran).Q.SQL.Text);
end;

//-------------------------------------------
//--- Nom   : FormKeyDown
//--- Objet : Scrute l'appuit d'une touche
//-------------------------------------------
procedure TOF_TBMULCOMPTA.FormKeyDown (Sender : TObject; var key:Word; Shift:TShiftState);
begin
 TFMul(Ecran).FormKeyDown(Sender, Key, Shift);
 case key of
  80,112 : if (Shift=[ssCtrl]) then
            Aff_TBStatCompta (TFMul(Ecran).Q.SQL.Text);
 end;
end;

Initialization
  registerclasses ( [ TOF_TBMULCOMPTA ] ) ;
end.

