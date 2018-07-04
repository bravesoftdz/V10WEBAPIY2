{***********UNITE*************************************************
Auteur  ...... : JL
Créé le ...... : 26/04/2004
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : FORMATIONCARRIERE ()
Mots clefs ... : TOF;FORMATIONCARRIERE
*****************************************************************
}
Unit UTofPGFormationCarriere;

Interface

Uses StdCtrls,Controls,Classes,
{$IFNDEF EAGLCLIENT}
     db,{$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}Fe_Main,
{$ELSE}
      MaineAgl,
{$ENDIF}
     forms,sysutils,ComCtrls,HCtrls,HEnt1,HMsgBox,UTOF,uTableFiltre,
     SaisieList,hTB97,LookUp,UTOB,ParamDat,HSysMenu;

Type
  TOF_PGFORMATIONCARRIERE = Class (TOF)
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    private
    TF : TTableFiltre;
    LeSalarie,LaNature : String;
    procedure StageElipsisClick(Sender : TObject);
    procedure DateElipsisclick(Sender :  TObject);
    procedure AjouterFormation(Sender : TObject);
  end ;

Implementation


procedure TOF_PGFORMATIONCARRIERE.OnUpdate ;
begin
  Inherited ;
end ;

procedure TOF_PGFORMATIONCARRIERE.OnLoad ;
begin
  Inherited ;
        SetControlText('XX_WHERE','EXISTS (SELECT PFO_ORDRE,PFO_CODESTAGE FROM FORMATIONS WHERE PFO_CODESTAGE =PSS_CODESTAGE AND PFO_ORDRE=PSS_ORDRE AND PFO_SALARIE="'+LeSalarie+'" AND PFO_MILLESIME=PSS_MILLESIME)');
        TFSaisieList(Ecran).BCherche.Click;
end ;


procedure TOF_PGFORMATIONCARRIERE.OnArgument (S : String ) ;
var BAjoutSession : TToolBarButton97;
    Utilisateur : String;
begin
  Inherited ;
        Utilisateur := V_PGI.UserSalarie;
        TF  :=  TFSaisieList(Ecran).LeFiltre;
        LeSalarie := Trim(ReadTokenPipe(S,';'));
        LaNature := Trim(ReadTokenPipe(S,';'));
        If LaNature = 'INITIALE' then SetControlText('PSS_NATUREFORM','004');
        If LaNature = 'EXTERNE' then SetControlText('PSS_NATUREFORM','002');
        If LaNature = 'INTERNE' then SetControlText('PSS_NATUREFORM','001');
        BAjoutSession := TToolBarButton97(GetControl('BAJOUTSESSION'));
        if BAjoutSession <> Nil then BAjoutSession.OnClick := AjouterFormation;
        TF.WhereTable := 'WHERE PFO_CODESTAGE=:PSS_CODESTAGE AND PFO_ORDRE=:PSS_ORDRE AND PFO_MILLESIME=:PSS_MILLESIME AND PFO_SALARIE="'+LeSalarie+'"';
end ;

procedure TOF_PGFORMATIONCARRIERE.StageElipsisClick(Sender : TObject);
begin
        LookupList(THEdit(Sender),'Liste des stages','STAGE','PST_CODESTAGE','PST_LIBELLE,PST_MILLESIME','','', True,-1);
end;

procedure TOF_PGFORMATIONCARRIERE.DateElipsisclick(Sender : TObject);
var key : char;
begin
    key  :=  '*';
    ParamDate (Ecran, Sender, Key);
end;

procedure TOF_PGFORMATIONCARRIERE.AjouterFormation(Sender : TObject);
begin
        TF.DisableTOM;
        AGLLanceFiche('PAY','MUL_SESSIONSTAGE','','','CARRIERE;'+LaNature+';'+LeSalarie);
        TF.EnableTOM;
        TF.RefreshEntete;
end;

Initialization
  registerclasses ( [ TOF_PGFORMATIONCARRIERE ] ) ;
end.
