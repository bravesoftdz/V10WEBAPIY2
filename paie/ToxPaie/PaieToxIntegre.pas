{***********UNITE*************************************************
Auteur  ...... : Dev Paie
Créé le ...... : 09/04/2003
Modifié le ... :   /  /
Description .. : Intégration des fichiers TOX
Suite ........ : Module de contrôle et d'intégration des données
Mots clefs ... : TOX;CONTROLE;INTEGRATION
*****************************************************************}
unit PaieToxIntegre;
interface

uses DBTables,
     Hctrls,
     SysUtils,
     uTob,
     StdCtrls,
     Dialogs,
     PaieToxMessage,
     ParamSoc ;

//////////////////////////////////////////////////////////////////////////////
// Liste des fonctions permettant de contrôler ou d'intégrer un enregistrement
///////////////////////////////////////////////////////////////////////////////
function PaieImportTOBPaieParam    (TPST : TOB; Integre, IntegrationRapide : boolean ; NomTable : String) : integer ;
function PaieImportTOBTOX          (TPST : TOB; Integre, IntegrationRapide : boolean ; NomTable : String)  : integer ;
function PaieImportTOBCodepost     (TPST : TOB; Integre, IntegrationRapide : boolean ) : integer ;
function PaieImportTOBChoixcod     (TPST : TOB; Integre, IntegrationRapide : boolean ) : integer ;
function PaieImportTOBChoixExt     (TPST : TOB; Integre, IntegrationRapide : boolean ) : integer ;
function PaieImportTOBEtabliss     (TPST : TOB; Integre, IntegrationRapide : boolean ) : integer ;
function PaieImportTOBLiensOLE     (TPST : TOB; Integre, IntegrationRapide : boolean ) : integer ;
function PaieImportTOBModeles      (TPST : TOB; Integre, IntegrationRapide : boolean ) : integer ;
function PaieImportTOBModedata     (TPST : TOB; Integre, IntegrationRapide : boolean ) : integer ;
function PaieImportTOBPays         (TPST : TOB; Integre, IntegrationRapide : boolean ) : integer ;
function PaieImportTOBTradDico     (TPST : TOB; Integre, IntegrationRapide : boolean ) : integer ;
function PaieImportTOBTradTablette (TPST : TOB; Integre, IntegrationRapide : boolean ) : integer ;
function PaieImportTOBTraduc       (TPST : TOB; Integre, IntegrationRapide : boolean ) : integer ;
function PaieImportTOBUserGrp      (TPST : TOB; Integre, IntegrationRapide : boolean ) : integer ;
function PaieImportTOBUtilisat     (TPST : TOB; Integre, IntegrationRapide : boolean ) : integer ;

function PaieToxControleInfopaie   (NomTable : string; var ToxInfo : TOB; Integre, IntegrationRapide : Boolean) : integer ;
Function PaieToxRendTablePaie (NomTable : String ) : TOB ;
Procedure PaieToxChargeInfoPaie ();
Procedure PaieToxLibereInfoPaie ();

var DernierTypeChoixcod : string ;
    DernierTypeChoixext : string ;
    MesInfosPaie,MesInfosTox,TT     : TOB ;
implementation

{***********UNITE*************************************************
Auteur  ...... : Dev Paie
Créé le ...... : 09/04/2003
Modifié le ... : 09/04/2003
Description .. : Intégration des fichiers TOX
Suite ........ : fonction générique de reconnaissance et des tables
Suite ........ : de la paie
Suite ........ :
Mots clefs ... : TOX;CONTROLE;INTEGRATION
*****************************************************************}

function PaieImportTOBPaieParam    (TPST : TOB; Integre,  IntegrationRapide : boolean ; NomTable : String) : integer ;
var MaTob        : TOB ;
begin
  Result:=IntegOK;

  if not (Integre) then
  begin
    if not (IntegrationRapide) then 	// PHASE DE CONTROLE
    begin
      // Contrôle(s)
    end ;
  end
  else
  begin 							// PHASE D'INTEGRATION
    ////////////////////////////////////////////////////////////////
    // Création d'une TOB contenant l'enregistrement courant
    ////////////////////////////////////////////////////////////////
    MaTob := TOB.Create (NomTable, nil, -1);
    MaTob.Dupliquer (TPST, False, True, True);
    MaTob.SetAllModifie (True);

   	if Not MaTob.InsertOrUpdateDB (FALSE) then
    begin
      Result:=ErrUpdateDB;
    end;
    MaTob.free;
  end;
end;

{***********UNITE*************************************************
Auteur  ...... : Dev Paie
Créé le ...... : 09/04/2003
Modifié le ... : 09/04/2003
Description .. : Intégration des fichiers TOX
Suite ........ : fonction générique de reconnaissance des tables concernant la TOX
Suite ........ :
Suite ........ :
Mots clefs ... : TOX;CONTROLE;INTEGRATION
*****************************************************************}

function PaieImportTOBTOX  (TPST : TOB; Integre, IntegrationRapide : boolean; NomTable : String)  : integer ;
var MaTob        : TOB ;
begin
  Result:=IntegOK;

  if not (Integre) then
  begin
    if not (IntegrationRapide) then 	// PHASE DE CONTROLE
    begin
      // Contrôle(s)
    end ;
  end
  else
  begin 							// PHASE D'INTEGRATION
    ////////////////////////////////////////////////////////////////
    // Création d'une TOB contenant l'enregistrement courant
    ////////////////////////////////////////////////////////////////
    MaTob := TOB.Create (NomTable, nil, -1);
    MaTob.Dupliquer (TPST, False, True, True);
    MaTob.SetAllModifie (True);

   	if Not MaTob.InsertOrUpdateDB (FALSE) then
    begin
      Result:=ErrUpdateDB;
    end;
    MaTob.free;
  end;
end;


{***********A.G.L.Privé.*****************************************
Auteur  ...... : Dev Paie
Créé le ...... : 09/04/2003
Modifié le ... :   /  /
Description .. : Contrôles/Import TOB CHOIXCOD
Mots clefs ... : TOX
*****************************************************************}
function PaieImportTOBChoixcod (TPST : TOB; Integre,  IntegrationRapide : boolean ) : integer ;
var Tob_Choixcod : TOB ;
  	CCType       : string ;
    CCCode       : string ;
    SQL          : string ;
    Q            : TQUERY ;
begin

  Result:=IntegOK;

  if not (Integre) then
  begin
    if not (IntegrationRapide) then 	// PHASE DE CONTROLE
    begin
      // Contrôle(s)
    end ;
  end
  else begin 							// PHASE D'INTEGRATION
    ////////////////////////////////////////////////////////////////
    // Création d'une TOB contenant l'enregistrement courant
    ////////////////////////////////////////////////////////////////
    Tob_ChoixCod:= TOB.Create ('CHOIXCOD', nil, -1);
    Tob_ChoixCod.Dupliquer (TPST, False, True, True);
    Tob_ChoixCod.SetAllModifie(True);

    CCType := Tob_ChoixCod.GetValue ('CC_TYPE');
    CCCode := Tob_ChoixCod.GetValue ('CC_CODE');

    ////////////////////////////////////////////////////////////////
    // MODIF LM 28112002 - Suppression des enregistrements CHOIXCOD
    ////////////////////////////////////////////////////////////////
    if CCType <> DernierTypeChoixcod then
    begin
      SQL:='DELETE From CHOIXCOD WHERE CC_TYPE="'+CCType+'"';
      ExecuteSQL(SQL) ;
      // Sauvegarde du dernier type CHOIXCOD supprimé
      DernierTypeChoixcod := CCType ;
    end;

    SQL:='Select CC_LIBELLE From CHOIXCOD WHERE CC_TYPE="'+CCType+'" AND CC_CODE="'+CCCode+'"';
    Q:=OpenSQL(SQL,True) ;
    if Not Q.EOF then
    begin
    	if Not Tob_ChoixCod.UpdateDB (FALSE) then
      begin
        Result:=ErrUpdateDB;
      end;
    end
    else begin
    	if Not Tob_ChoixCod.InsertDB (Nil) then
      begin
        Result:=ErrInsertDB;
      end;
    end;
    Ferme (Q);
    Tob_ChoixCod.free;
  end;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : Dev Paie
Créé le ...... : 09/04/2003
Modifié le ... :   /  /
Description .. : Contrôles/Import TOB CHOIXEXT
Mots clefs ... : TOX
*****************************************************************}
function PaieImportTOBChoixExt (TPST : TOB; Integre,  IntegrationRapide : boolean ) : integer ;
var Tob_ChoixExt : TOB ;
		YXType       : string ;
    YXCode       : string ;
    SQL          : string ;
    Q            : TQUERY ;
begin

  Result:=IntegOK;

	if not (Integre) then
  begin
    if not (IntegrationRapide) then 	// PHASE DE CONTROLE
    begin
      // Contrôle(s)
    end ;
  end
  else begin 							// PHASE D'INTEGRATION
    ////////////////////////////////////////////////////////////////
    // Création d'une TOB contenant l'enregistrement courant
    ////////////////////////////////////////////////////////////////
    Tob_ChoixExt:= TOB.Create ('CHOIXEXT', nil, -1);
    Tob_ChoixExt.Dupliquer (TPST, False, True, True);
    Tob_ChoixExt.SetAllModifie(True);

    YXType := Tob_ChoixExt.GetValue ('YX_TYPE');
    YXCode := Tob_ChoixExt.GetValue ('YX_CODE');

    //////////////////////////////////////////////////////////////////
    // MODIF LM 28/11/2002 - Suppression des enregistrements CHOIXEXT
    //////////////////////////////////////////////////////////////////
    if YXType <> DernierTypeChoixext then
    begin
      SQL:='DELETE From CHOIXEXT WHERE YX_TYPE="'+YXType+'"';
      ExecuteSQL(SQL) ;
      // Sauvegarde du dernier type CHOIXCOD supprimé
      DernierTypeChoixext := YXType ;
    end;

    SQL:='Select YX_LIBELLE From CHOIXEXT WHERE YX_TYPE="'+YXType+'" AND YX_CODE="'+YXCode+'"';
    Q:=OpenSQL(SQL,True) ;
    if Not Q.EOF then
    begin
    	if Not Tob_ChoixExt.UpdateDB (FALSE) then
      begin
        Result:=ErrUpdateDB;
      end;
    end
    else begin
    	if Not Tob_ChoixExt.InsertDB (Nil) then
      begin
        Result:=ErrInsertDB;
      end;
    end;
    Ferme (Q);
    Tob_ChoixExt.free;
  end;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : Dev Paie
Créé le ...... : 09/04/2003
Modifié le ... :   /  /
Description .. : Contrôles/Import TOB CODEPOST
Mots clefs ... : TOX
*****************************************************************}
function PaieImportTOBCodepost (TPST : TOB; Integre,  IntegrationRapide : boolean ) : integer ;
var Tob_Codepost : TOB ;
		Postal       : string ;
    Ville        : string ;
    SQL          : string ;
    Q            : TQUERY ;
begin

  Result:=IntegOK;

	if not (Integre and IntegrationRapide) then 	// PHASE DE CONTROLE
  begin
  	// Contrôle(s)
  end
  else begin 							// PHASE D'INTEGRATION
    ////////////////////////////////////////////////////////////////
    // Création d'une TOB contenant l'enregistrement courant
    ////////////////////////////////////////////////////////////////
    Tob_Codepost:= TOB.Create ('CODEPOST', nil, -1);
    Tob_Codepost.Dupliquer (TPST, False, True, True);
    Tob_Codepost.SetAllModifie(True);

    Postal := Tob_Codepost.GetValue ('O_CODEPOSTAL');
    Ville  := Tob_Codepost.GetValue ('O_VILLE');
    SQL:='Select O_PAYS From CODEPOST WHERE O_CODEPOSTAL"'+Postal+'" AND O_VILLE="'+Ville+'"';
    Q:=OpenSQL(SQL,True) ;
    if Not Q.EOF then
    begin
    	if Not Tob_Codepost.UpdateDB (FALSE) then
      begin
        Result:=ErrUpdateDB;
      end;
    end
    else begin
    	if Not Tob_Codepost.InsertDB (Nil) then
      begin
        Result:=ErrInsertDB;
      end;
    end;
    Ferme (Q);
    Tob_Codepost.free;
  end;
end;
{***********A.G.L.Privé.*****************************************
Auteur  ...... : Dev Paie
Créé le ...... : 09/04/2003
Modifié le ... :   /  /
Description .. : Contrôles/Import TOB ETABLISS
Mots clefs ... : TOX
*****************************************************************}
function PaieImportTOBEtabliss (TPST : TOB; Integre,  IntegrationRapide : boolean ) : integer ;
var Tob_Etab : TOB ;
		CodeEtab : string ;
    StRech   : string ;
    SQL      : string ;
    Q        : TQUERY ;
begin

	Result:=IntegOK;

	if not (Integre) then
  begin
    if not (IntegrationRapide) then 	// PHASE DE CONTROLE
    begin
      /////////////////////////////////////////////////////////////////////////
      // Pays
      /////////////////////////////////////////////////////////////////////////
      StRech := TPST.GetValue ('ET_PAYS');
      if (StRech <> '') then
      begin
        SQL:='Select PY_LIBELLE From PAYS WHERE PY_PAYS="'+StRech+'"';
        Q:=OpenSQL(SQL,True) ;
        if Q.EOF then result := ErrPays;
        Ferme(Q);
      end;
      /////////////////////////////////////////////////////////////////////////
      // Langue
      /////////////////////////////////////////////////////////////////////////
      if (result=IntegOK) then
      begin
        StRech := TPST.GetValue ('ET_LANGUE');
        if (StRech <> '') then
        begin
        SQL:='Select CC_LIBELLE From CHOIXCOD WHERE CC_TYPE="LGU" AND CC_CODE="'+StRech+'"';
          Q:=OpenSQL(SQL,True) ;
          if Q.EOF then result := ErrLangue;
          Ferme(Q);
        end;
      end;
    end ;
  end
  else begin 							// PHASE D'INTEGRATION
  	////////////////////////////////////////////////////////////////
    // Création d'une TOB contenant l'enregistrement courant
	  ////////////////////////////////////////////////////////////////
    Tob_Etab:= TOB.Create ('ETABLISS', nil, -1);
    Tob_Etab.Dupliquer (TPST, False, True, True);
    Tob_Etab.SetAllModifie(True);

    //
    // doit-on conserver la coche "Géré sur site" ?
    //
    CodeEtab := Tob_Etab.GetValue ('ET_ETABLISSEMENT');
    SQL:='Select ET_ETABLISSEMENT From ETABLISS WHERE ET_ETABLISSEMENT="'+CodeEtab+'"';
    Q:=OpenSQL(SQL,True) ;
    if Not Q.EOF then
    begin
    	if Not Tob_Etab.UpdateDB (FALSE) then
      begin
        Result:=ErrUpdateDB;
      end;
    end
    else begin
    	if Not Tob_Etab.InsertDB (Nil) then
      begin
        Result:=ErrInsertDB;
      end;
    end;
    Ferme (Q);
    Tob_Etab.free;
  end;
end;
{***********A.G.L.Privé.*****************************************
Auteur  ...... : Dev Paie
Créé le ...... : 09/04/2003
Modifié le ... :   /  /
Description .. : Contrôles/Import TOB LIENSOLE
Mots clefs ... : TOX
*****************************************************************}
function PaieImportTOBLiensOLE (TPST : TOB; Integre,  IntegrationRapide : boolean ) : integer ;
var Tob_LiensOLE :  TOB   ;
		Table        : string ;
		Ident   		 : string ;
    Rang         : integer;
    SQL          : string ;
    Q            : TQUERY ;
begin

	Result:=IntegOK;

	if not (Integre) then
  begin
    if not (IntegrationRapide) then 	// PHASE DE CONTROLE
    begin
      // Contrôle(s)
    end ;
  end
  else begin 							// PHASE D'INTEGRATION
	  ////////////////////////////////////////////////////////////////
    // Création d'une TOB contenant l'enregistrement courant
	  ////////////////////////////////////////////////////////////////
    Tob_LiensOLE := TOB.Create ('LIENOLE', nil, -1);
    Tob_LiensOLE.Dupliquer (TPST, False, True, True);
    Tob_LiensOLE.SetAllModifie(True);

    Table  := Tob_LiensOLE.GetValue ('LO_TABLEBLOB');
    Ident  := Tob_LiensOLE.GetValue ('LO_IDENTIFIANT');
    Rang   := Tob_LiensOLE.GetValue ('LO_RANGBLOB');

    SQL:='Select LO_EMPLOIBLOB From LIENSOLE WHERE LO_TABLEBLOB="'+Table+'" AND LO_IDENTIFIANT="'+Ident+'" AND LO_RANGBLOB="'+IntToStr(Rang)+'"';
    Q:=OpenSQL(SQL,True) ;
    if Not Q.EOF then
    begin
    	if Not Tob_LiensOLE.UpdateDB (FALSE) then
      begin
        Result:=ErrUpdateDB;
      end;
    end
    else begin
    	if Not Tob_LiensOLE.InsertDB (Nil) then
      begin
        Result:=ErrInsertDB;
      end;
    end;
    Ferme (Q);
    Tob_LiensOLE.free;
  end;
end;
{***********A.G.L.Privé.*****************************************
Auteur  ...... : Dev Paie
Créé le ...... : 09/04/2003
Modifié le ... :   /  /    
Description .. : Contrôles/Import TOB MODELES
Mots clefs ... : TOX
*****************************************************************}
function PaieImportTOBModeles (TPST : TOB; Integre,  IntegrationRapide : boolean ) : integer ;
var Tob_Modeles :  TOB ;
    StRech      : string ;
		TypeModele  : string ;
		Nature  		: string ;
    Code        : string ;
    Langue      : string ;
    SQL         : string ;
    Q           : TQUERY ;
begin

	Result:=IntegOK;

	if not (Integre) then
  begin
    if not (IntegrationRapide) then 	// PHASE DE CONTROLE
    begin
      /////////////////////////////////////////////////////////////////////////
      // Langue
      /////////////////////////////////////////////////////////////////////////
      if (result=IntegOK) then
      begin
        StRech := TPST.GetValue ('MO_LANGUE');
        if (StRech <> '') then
        begin
          SQL:='Select CC_LIBELLE From CHOIXCOD WHERE CC_TYPE="LGU" AND CC_CODE="'+StRech+'"';
          Q:=OpenSQL(SQL,True) ;
          if Q.EOF then Result:=ErrLangue;
          Ferme (Q);
        end;
      end;
    end ;
  end
  else begin 							// PHASE D'INTEGRATION
	  ////////////////////////////////////////////////////////////////
    // Création d'une TOB contenant l'enregistrement courant
	  ////////////////////////////////////////////////////////////////
    Tob_Modeles:= TOB.Create ('MODELES', nil, -1);
    Tob_Modeles.Dupliquer (TPST, False, True, True);
    Tob_Modeles.SetAllModifie(True);

    TypeModele := Tob_Modeles.GetValue ('MO_TYPE');
    Nature     := Tob_Modeles.GetValue ('MO_NATURE');
    Code       := Tob_Modeles.GetValue ('MO_CODE');
    Langue     := Tob_Modeles.GetValue ('MO_LANGUE');

    SQL:='Select MO_LIBELLE From MODELES WHERE MO_TYPE="'+TypeModele+'" AND MO_NATURE="'+Nature+'" AND MO_CODE="'+Code+'" AND MO_LANGUE="'+Langue+'"';
    Q:=OpenSQL(SQL,True) ;
    if Not Q.EOF then
    begin
    	if Not Tob_Modeles.UpdateDB (FALSE) then
      begin
        Result:=ErrUpdateDB;
      end;
    end
    else begin
    	if Not Tob_Modeles.InsertDB (Nil) then
      begin
        Result:=ErrInsertDB;
      end;
    end;
    Ferme (Q);
    Tob_Modeles.free;
  end;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : Dev Paie
Créé le ...... : 09/04/2003
Modifié le ... :   /  /
Description .. : Contrôles/Import TOB MODEDATA
Mots clefs ... : TOX
*****************************************************************}
function PaieImportTOBModeData (TPST : TOB; Integre,  IntegrationRapide : boolean ) : integer ;
var Tob_ModeData :  TOB ;
    Clef         : string ;
    SQL          : string ;
    Q            : TQUERY ;
begin

	Result:=IntegOK;

	if not (Integre) then
  begin
    if not (IntegrationRapide) then 	// PHASE DE CONTROLE
    begin
      /////////////////////////////////////////////////////////////////////////
      // Contrôles ....
      /////////////////////////////////////////////////////////////////////////
    end ;
  end
  else begin 							// PHASE D'INTEGRATION
	  ////////////////////////////////////////////////////////////////
    // Création d'une TOB contenant l'enregistrement courant
	  ////////////////////////////////////////////////////////////////
    Tob_ModeData:= TOB.Create ('MODEDATA', nil, -1);
    Tob_ModeData.Dupliquer (TPST, False, True, True);
    Tob_ModeData.SetAllModifie(True);

    Clef  := Tob_ModeData.GetValue ('MD_CLE');

    SQL:='Select MD_CLE From MODEDATA WHERE MD_CLE="'+Clef+'"';
    Q:=OpenSQL(SQL,True) ;
    if Not Q.EOF then
    begin
    	if Not Tob_ModeData.UpdateDB (FALSE) then
      begin
        Result:=ErrUpdateDB;
      end;
    end
    else begin
    	if Not Tob_ModeData.InsertDB (Nil) then
      begin
        Result:=ErrInsertDB;
      end;
    end;
    Ferme (Q);
    Tob_ModeData.free;
  end;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : Dev Paie
Créé le ...... : 09/04/2003
Modifié le ... :   /  /
Description .. : Contrôles/Import TOB PAYS
Mots clefs ... : TOX
*****************************************************************}
function PaieImportTOBPays (TPST : TOB; Integre,  IntegrationRapide : boolean ) : integer ;
var Tob_Pays : TOB    ;
		CodePays : string ;
    StRech   : string ;
    SQL      : string ;
    Q        : TQUERY ;
begin

  Result:=IntegOK;

	if not (Integre) then
  begin
    if not (IntegrationRapide) then 	// PHASE DE CONTROLE
    begin
      /////////////////////////////////////////////////////////////////////////
      // Langue
      /////////////////////////////////////////////////////////////////////////
      if (result=IntegOK) then
      begin
        StRech := TPST.GetValue ('PY_LANGUE');
        if (StRech <> '') then
        begin
          SQL:='Select CC_LIBELLE From CHOIXCOD WHERE CC_TYPE="LGU" AND CC_CODE="'+StRech+'"';
          Q:=OpenSQL(SQL,True) ;
          if Q.EOF then Result:=ErrLangue;
          Ferme (Q);
        end;
      end;
    end ;
  end
  else begin 							// PHASE D'INTEGRATION
    ////////////////////////////////////////////////////////////////
    // Création d'une TOB contenant l'enregistrement courant
	  ////////////////////////////////////////////////////////////////
    Tob_Pays:= TOB.Create ('PAYS', nil, -1);
    Tob_Pays.Dupliquer (TPST, False, True, True);
    Tob_Pays.SetAllModifie(True);

    CodePays := Tob_Pays.GetValue ('PY_PAYS');
    SQL:='Select PY_LIBELLE From PAYS WHERE PY_PAYS="'+CodePays+'"';
    Q:=OpenSQL(SQL,True) ;
    if Not Q.EOF then
    begin
    	if Not Tob_Pays.UpdateDB (FALSE) then
      begin
        Result:=ErrUpdateDB;
      end;
    end
    else begin
    	if Not Tob_Pays.InsertDB (Nil) then
      begin
        Result:=ErrInsertDB;
      end;
    end;
    Ferme (Q);
    Tob_Pays.free;
  end;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : Dev Paie
Créé le ...... : 09/04/2003
Modifié le ... :   /  /
Description .. : Contrôles/Import TOB TRADICCO
Mots clefs ... : TOX
*****************************************************************}
function PaieImportTOBTradDico (TPST : TOB; Integre, IntegrationRapide : boolean ) : integer ;
var SQL,CodeDico : string ;
    Tob_TradDico : TOB ;
begin

  Result:=IntegOK;

	if not (Integre) then
  begin
    if not (IntegrationRapide) then 	// PHASE DE CONTROLE
    begin
    // pas de contrôle
    end ;
  end
  else begin 							// PHASE D'INTEGRATION

    // Création d'une TOB contenant  l'enregistrement courant
    Tob_TradDico := TOB.Create ('TRADDICO', nil, -1);
    Tob_TradDico.Dupliquer (TPST, False, True, True);
    Tob_TradDico.SetAllModifie(True);

    CodeDico := Tob_TradDico.GetValue ('DX_FRA');

    SQL:='Select DX_FRA From TRADDICO WHERE DX_FRA="'+CodeDico+'"';
    if ExisteSQL(SQL) then
    	begin
      if Not Tob_TradDico.UpdateDB (FALSE) then Result:=ErrUpdateDB ;
      end
    else
    	if Not Tob_TradDico.InsertDB (Nil)   then Result:=ErrInsertDB ;
    Tob_TradDico.free;
  end;
end ;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : Dev Paie
Créé le ...... : 09/04/2003
Modifié le ... :   /  /
Description .. : Contrôles/Import TOB TRADTABLETTE
Mots clefs ... : TOX
*****************************************************************}
function PaieImportTOBTradTablette (TPST : TOB; Integre, IntegrationRapide : boolean ) : integer ;
var sLangue, sTypeTablette, sCodeTablette, StRech : string ;
    Tob_TradTablette : TOB ;
begin

  Result:=IntegOK;

	if not (Integre) then
  begin
    if not (IntegrationRapide) then 	// PHASE DE CONTROLE
    begin
      //
      // Langue
      //
      if (result=IntegOK) then
      begin
        StRech := TPST.GetValue ('YTT_LANGUE');
        if (StRech <> '') then
          if not ExisteSQL('Select CC_LIBELLE From CHOIXCOD WHERE CC_TYPE="LGU" AND CC_CODE="'+StRech+'"') then
            result := ErrLangue;
      end;
      //
      // Code
      //
      if (result=IntegOK) then
      begin
        StRech := TPST.GetValue ('YTT_CODE');
        if (StRech <> '') then
          if not ExisteSQL('Select DO_COMBO From DECOMBOS WHERE DO_COMBO="'+StRech+'"') then
            result := ErrCodeTablette;
      end;
    end ;
  end
  else begin 							// PHASE D'INTEGRATION

    // Création d'une TOB contenant  l'enregistrement courant
    Tob_TradTablette := TOB.Create ('TRADTABLETTE', nil, -1);
    Tob_TradTablette.Dupliquer (TPST, False, True, True);
    Tob_TradTablette.SetAllModifie(True);

    sLangue       := Tob_TradTablette.GetValue ('YTT_LANGUE');
    sTypeTablette := Tob_TradTablette.GetValue ('YTT_TYPE');
    sCodeTablette := Tob_TradTablette.GetValue ('YTT_CODE');

    if ExisteSQL('Select YTT_LANGUE From TRADTABLETTE WHERE YTT_LANGUE="'+sLangue+'" AND YTT_TYPE="'+sTypeTablette+'" AND YTT_CODE="'+sCodeTablette+'"') then
    	begin if Not Tob_TradTablette.UpdateDB (FALSE) then Result:=ErrUpdateDB ; end
    else
    	begin if Not Tob_TradTablette.InsertDB (Nil)   then Result:=ErrInsertDB ; end ;

    Tob_TradTablette.free;
  end;

end ;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : Dev Paie
Créé le ...... : 09/04/2003
Modifié le ... :   /  /
Description .. : Contrôles/Import TOB TRADUC
Mots clefs ... : TOX
*****************************************************************}
function PaieImportTOBTraduc (TPST : TOB; Integre, IntegrationRapide : boolean ) : integer ;
var sLangue, sForme, StRech : string ;
    Tob_Traduc : TOB ;
begin

  Result:=IntegOK;

	if not (Integre) then
  begin
    if not (IntegrationRapide) then 	// PHASE DE CONTROLE
    begin
      //
      // Langue
      //
      if (result=IntegOK) then
      begin
        StRech := TPST.GetValue ('TR_LANGUE');
        if (StRech <> '') then
          if not ExisteSQL('Select CC_LIBELLE From CHOIXCOD WHERE CC_TYPE="LGU" AND CC_CODE="'+StRech+'"') then
            result := ErrLangue;
      end;
      //
      // Forme
      //
      if (result=IntegOK) then
      begin
        StRech := TPST.GetValue ('TR_FORME');
        if (StRech <> '') then
          if not ExisteSQL('Select DFM_FORME From FORMES WHERE DFM_TYPEFORME="GC" AND DFM_FORME="'+StRech+'"') then
            result := ErrForme;
      end;
    end ;
  end
  else begin 							// PHASE D'INTEGRATION

    // Création d'une TOB contenant l'enregistrement courant
    Tob_Traduc := TOB.Create ('TRADUC', nil, -1);
    Tob_Traduc.Dupliquer (TPST, False, True, True);
    Tob_Traduc.SetAllModifie(True);

    sLangue := Tob_Traduc.GetValue ('TR_LANGUE');
    sForme  := Tob_Traduc.GetValue ('TR_FORME');

    if ExisteSQL('Select TR_LANGUE From TRADUC WHERE TR_LANGUE="'+sLangue+'" AND TR_FORME="'+sForme+'"') then
      begin if Not Tob_Traduc.UpdateDB (FALSE) then Result:=ErrUpdateDB ; end
    else
      begin if Not Tob_Traduc.InsertDB (Nil)   then Result:=ErrInsertDB ; end ;

    Tob_Traduc.free;
  end;

end ;
{***********A.G.L.Privé.*****************************************
Auteur  ...... : Dev Paie
Créé le ...... : 09/04/2003
Modifié le ... :   /  /
Description .. : Contrôles/Import TOB USERGRP
Mots clefs ... : TOX
*****************************************************************}
function PaieImportTOBUserGrp (TPST : TOB; Integre,  IntegrationRapide : boolean ) : integer ;
var Tob_UserGrp : TOB    ;
		Groupe 			: string ;
    SQL   			: string ;
    Q     		 	: TQUERY ;
begin

	Result:=IntegOK;

	if not (Integre) then
  begin
    if not (IntegrationRapide) then 	// PHASE DE CONTROLE
    begin
      // Contrôle(s)
    end ;
  end
  else begin 							// PHASE D'INTEGRATION
	  ////////////////////////////////////////////////////////////////
    // Création d'une TOB contenant l'enregistrement courant
	  ////////////////////////////////////////////////////////////////
    Tob_UserGrp := TOB.Create ('USERGRP', nil, -1);
    Tob_UserGrp.Dupliquer (TPST, False, True, True);
    Tob_UserGrp.SetAllModifie(True);

    Groupe := Tob_UserGrp.GetValue ('UG_GROUPE');
    SQL:='Select UG_LIBELLE From USERGRP WHERE UG_GROUPE="'+Groupe+'"';
    Q:=OpenSQL(SQL,True) ;
    if Not Q.EOF then
    begin
    	if Not Tob_UserGrp.UpdateDB (FALSE) then
      begin
        Result:=ErrUpdateDB;
      end;
    end
    else begin
    	if Not Tob_UserGrp.InsertDB (Nil) then
      begin
        Result:=ErrInsertDB;
      end;
    end;
    Ferme (Q);
    Tob_UserGrp.free;
  end;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : D. Carret
Créé le ...... : 10/09/2002
Modifié le ... :   /  /
Description .. : Contrôles/Import TOB UTILISAT
Mots clefs ... : TOX
*****************************************************************}
function PaieImportTOBUtilisat (TPST : TOB; Integre,  IntegrationRapide : boolean ) : integer ;
var Tob_Utilisat   : TOB    ;
	Utilisat       : string ;
    SQL   		   : string ;
    Q     		   : TQUERY ;
begin

  Result:=IntegOK;

  if not (Integre) then
  begin
    if not (IntegrationRapide) then 	// PHASE DE CONTROLE
    begin
      /////////////////////////////////////////////////////////////////////////
      // Pas de contrôle
      /////////////////////////////////////////////////////////////////////////
    end ;
  end
  else begin 							// PHASE D'INTEGRATION
	////////////////////////////////////////////////////////////////
    // Création d'une TOB contenant l'enregistrement courant
	////////////////////////////////////////////////////////////////
    Tob_Utilisat := TOB.Create ('UTILISAT', nil, -1);
    Tob_Utilisat.Dupliquer (TPST, False, True, True);
    Tob_Utilisat.SetAllModifie(True);

    Utilisat := Tob_Utilisat.GetValue ('US_UTILISATEUR');
    SQL:='Select US_UTILISATEUR From UTILISAT WHERE US_UTILISATEUR="'+Utilisat+'"';
    Q:=OpenSQL(SQL,True) ;
    if Not Q.EOF then
    begin
    	if Not Tob_Utilisat.UpdateDB (FALSE) then
      begin
        Result:=ErrUpdateDB;
      end;
    end
    else begin
    	if Not Tob_Utilisat.InsertDB (Nil) then
      begin
        Result:=ErrInsertDB;
      end;
    end;
    Ferme (Q);
    Tob_Utilisat.Free;
  end;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : Dev Paie
Créé le ...... : 09/04/2003
Modifié le ... :   /  /
Description .. : Fonction générique qui chaîne sur les fonctions
Suite ........ : spécifiques aux tables traitées
Mots clefs ... : TOX
*****************************************************************}
function PaieToxControleInfoPaie (NomTable : string; var ToxInfo : TOB; Integre, IntegrationRapide : boolean) : integer ;
var T1 : TOB ;
begin
  if (NomTable = 'CHOIXCOD')	      then result := PaieImportTOBChoixCod     (ToxInfo, Integre, IntegrationRapide)
  else if (NomTable = 'CHOIXEXT')	  then result := PaieImportTOBChoixExt     (ToxInfo, Integre, IntegrationRapide)
  else if (NomTable = 'ETABLISS')	  then result := PaieImportTOBEtabliss     (ToxInfo, Integre, IntegrationRapide)
  else if (NomTable = 'LIENSOLE')   then result := PaieImportTOBLiensOLE     (ToxInfo, Integre, IntegrationRapide)
  else if (NomTable = 'MODELES')	  then result := PaieImportTOBModeles      (ToxInfo, Integre, IntegrationRapide)
  else if (NomTable = 'MODEDATA')	  then result := PaieImportTOBModeData     (ToxInfo, Integre, IntegrationRapide)
  else if (NomTable = 'PAYS')	      then result := PaieImportTOBPays         (ToxInfo, Integre, IntegrationRapide)
  else if (NomTable = 'UTILISAT')  	then result := PaieImportTOBUtilisat     (ToxInfo, Integre, IntegrationRapide)
  else if (Copy (NomTable,1,4) = 'STOX') then result := PaieImportTOBTOX     (ToxInfo, Integre, IntegrationRapide, Nomtable)
  else
  begin
    T1 := PaieToxRendTablePaie (NomTable) ;
    if Assigned (T1) then result := PaieImportTOBPaieParam (ToxInfo, Integre, IntegrationRapide, NomTable )
       else   result := TableNonReconnu ;
  end;
end;

Procedure PaieToxChargeInfoPaie ();
var SQL   		   : string ;
    Q     		   : TQUERY ;
begin
  MesInfosPaie := TOB.Create ('DETABLES',NIL,-1) ;
  SQL := 'SELECT DT_NOMTABLE,DT_CLE1 FROM DETABLES WHERE DT_DOMAINE = "P" ';
  Q:=OPENSQL(SQL,TRUE) ;
  MesInfosPaie.LoadDetailDB ('DETABLES','', '', Q, FALSE) ;
  Ferme (Q);
  MesInfosTox := TOB.Create ('DETABLES',NIL,-1) ;
  SQL := 'SELECT DT_NOMTABLE,DT_CLE1 FROM DETABLES WHERE DT_NOMTABLE LIKE "STOX%" ';
  Q:=OPENSQL(SQL,TRUE) ;
  MesInfosTox.LoadDetailDB ('DETABLES','', '', Q, FALSE) ;
  Ferme (Q);
end;

Function PaieToxRendTablePaie (NomTable : String ) : TOB ;
begin
  Result := NIL ;
  if NOT Assigned (MesInfosPaie) then PaieToxChargeInfoPaie ;
  if (Copy (NomTable, 1, 4) <> 'STOX') then
  begin
    if (NOT Assigned (TT)) OR  (TT.GetValue ('DT_NOMTABLE') <> NomTable) then
      TT := MesInfosPaie.FindFirst (['DT_NOMTABLE'],[NomTable],FALSE) ;
    if Assigned (TT) AND (TT.GetValue ('DT_NOMTABLE') = NomTable) then result := TT ;
  end
  else
  begin
    if (NOT Assigned (TT)) OR  (TT.GetValue ('DT_NOMTABLE') <> NomTable) then
      TT := MesInfosTox.FindFirst (['DT_NOMTABLE'],[NomTable],FALSE) ;
    if Assigned (TT) AND (TT.GetValue ('DT_NOMTABLE') = NomTable) then result := TT ;
  end;
end;

Procedure PaieToxLibereInfoPaie ();
begin
  if Assigned (MesInfosPaie) then FreeAndNil(MesInfosPaie) ;
  if Assigned (MesInfosTox) then FreeAndNil(MesInfosTox) ;
end;

end.
