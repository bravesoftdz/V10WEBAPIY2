{***********UNITE*************************************************
Auteur  ...... : 
Créé le ...... : 03/09/2002
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : PGMULPLAFONDFORM ()
Mots clefs ... : TOF;PGMULPLAFONDFORM
*****************************************************************
---- JL 20/03/2006 modification clé annuaire ----
---- JL 17/10/2006 Modification contrôle des exercices de formations -----
PT1  17/05/2007 V_720 FL FQ 11532 Changement de la liste affichée en cas de gestion des frais par population
PT2  19/09/2007 V_80  FL Rafraîchissement de l'écran suite à la création d'un élément
}
Unit UtofPgMul_PlafondForm;

Interface

Uses StdCtrls,Controls,Classes,
{$IFNDEF EAGLCLIENT}
     DBGrids, HDB,Mul,Fiche,db,{$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}FE_Main,
{$ELSE}
     UtileAGL,MaineAgl,emul,
{$ENDIF}
     forms,sysutils,ComCtrls,HCtrls,HEnt1,HMsgBox,UTOF,Hqry,HTB97,UTob,PGOutilsFormation,PGPopulOutils,EntPaie ; //PT1

Type
  TOF_PGMULPLAFONDFORM = Class (TOF)
    procedure OnArgument (S : String ) ; override ;
    private
    procedure GrilleDblClick (Sender : TObject);
    procedure CreerPlafond (Sender : TObject);
  end ;

Implementation

procedure TOF_PGMULPLAFONDFORM.OnArgument (S : String ) ;
Var Q:TQuery;
    Millesime:String;
    Combo:THValComboBox;
    {$IFNDEF EAGLCLIENT}
    Liste:THDBGrid;
    {$ELSE}
    Liste:THGrid;
    {$ENDIF}
    BOuv:TToolBarButton97;
begin
  Inherited ;
        Millesime:=RendMillesimePrevisionnel;
        Combo:=THValComboBox(GetControl('PFP_MILLESIME'));
        Combo.Value:=Millesime;
        TFMul(Ecran).Caption:='Liste des plafonds '+Millesime+' pour les organismes de formation';
        UpdateCaption(TFMul(Ecran));
        {$IFNDEF EAGLCLIENT}
        Liste:=THDBGrid(GetControl('FLISTE'));
        {$ELSE}
        Liste:=THGrid(GetControl('FLISTE'));
        {$ENDIF}
        if Liste <> NIL then Liste.OnDblClick := GrilleDblClick;
        BOuv := TToolbarButton97 (getcontrol ('BInsert'));
        if Bouv <> NIL then  BOuv.OnClick := CreerPlafond;
        //PT1 - Début
        // Dans le cas de la gestion des plafonds avec population, il faut changer de liste
        If (VH_Paie.PGForGestPlafByPop) Then
        Begin
               TFMul(Ecran).SetDBListe('PGPLAFONDFORMPOP');
               SetControlVisible('PFP_POPULATION', True);
               SetControlVisible('TPFP_POPULATION', True);
               SetControlText('XX_WHERE','PFP_POPULATION<>""');
               // Mise à jour de la liste de la combo avec les populations actives
               SetControlProperty('PFP_POPULATION', 'Plus', ' AND PPC_PREDEFINI="'+GetPredefiniPopulation(TYP_POPUL_FORM_PREV)+'"');
        End
        Else
               SetControlText('XX_WHERE','PFP_POPULATION=""');
        //PT1 - Fin
end ;

procedure TOF_PGMULPLAFONDFORM.GrilleDblClick (Sender : TObject);
var Q_Mul:THQuery ;
    St : String;
    {$IFNDEF EAGLCLIENT}
    Liste:THDBGrid;
    {$ELSE}
    Liste:THGrid;
    {$ENDIF}
begin
        {$IFNDEF EAGLCLIENT}
        Liste:=THDBGrid(GetControl('FLISTE'));
        {$ELSE}
        Liste:=THGrid(GetControl('FLISTE'));
        {$ENDIF}
        {$IFDEF EAGLCLIENT}
        TFmul(Ecran).Q.TQ.Seek(Liste.Row-1) ;
        {$ENDIF}
        Q_Mul:=THQuery(Ecran.FindComponent('Q'));
        St := Q_Mul.FindField('PFP_MILLESIME').AsString+';'+Q_Mul.FindField('PFP_ORGCOLLECTGU').AsString+
        ';'+Q_Mul.FindField('PFP_FRAISSALFOR').AsString+';'+IntToStr(Q_Mul.FindField('PFP_ORDRE').AsInteger);
        AGLLanceFiche ('PAY', 'PLAFONDFORM', '',St, 'ACTION=MODIFICATION');
end;

procedure TOF_PGMULPLAFONDFORM.CreerPlafond (Sender : TObject);
var St : String;
begin
     St := GetControlText('PFP_MILLESIME');
     AGLLanceFiche ('PAY', 'PLAFONDFORM', '','',St+';ACTION=CREATION');
     TFMul(Ecran).BChercheClick(TFMul(Ecran).BCherche); //PT1
end;


Initialization
  registerclasses ( [ TOF_PGMULPLAFONDFORM ] ) ;
end.

