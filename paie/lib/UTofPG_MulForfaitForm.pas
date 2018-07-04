{***********UNITE*************************************************
Auteur  ...... : 
Créé le ...... : 06/08/2002
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : PGMULFORFAITFORM ()
Mots clefs ... : TOF;PGMULFORFAITFORM
*****************************************************************
PT1 24/11/2003 JL V_50 Ajout Q.TQ.Seek en CWAS
---- JL 17/10/2006 Modification contrôle des exercices de formations -----
PT2 14/05/2007 V_720 FL FQ 13567 Gestion des populations
}
Unit UTofPG_MulForfaitForm;

Interface

Uses Controls,Classes,
{$IFNDEF EAGLCLIENT}
     db,{$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}HDB,Mul,Fe_Main,
{$ELSE}
     Emul,UTOB,MaineAgl,
{$ENDIF}
     HCtrls,UTOF,Hqry,PGOutilsFormation, EntPaie, PGPopulOutils; //PT2

Type
  TOF_PGMULFORFAITFORM = Class (TOF)
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    private
    Q_Mul      : THQuery ;
    procedure GrilleDblClick(Sender : TObject);
    procedure CreateForfait(Sender : TObject);
  end ;

Implementation

procedure TOF_PGMULFORFAITFORM.OnUpdate ;
begin
  Inherited ;
end ;

procedure TOF_PGMULFORFAITFORM.OnLoad ;
begin
  Inherited ;
end ;

procedure TOF_PGMULFORFAITFORM.OnArgument (S : String ) ;
var Millesime : String;
    {$IFNDEF EAGLCLIENT}
    Grille : THDBGrid;
    {$ELSE}
    Grille : THGrid;
    {$ENDIF}
begin
  Inherited ;
        Millesime := RendMillesimePrevisionnel;  
        SetControltext('PFF_MILLESIME',Millesime);
        TFMul(Ecran).Caption := 'Liste des forfaits pour le prévisionnel '+Millesime;
        UpdateCaption(TFMul(Ecran)) ;
        Q_Mul := THQuery(Ecran.FindComponent('Q'));
        {$IFNDEF EAGLCLIENT}
        Grille := THDBGrid (GetControl ('Fliste'));
        {$ELSE}
        Grille := THGrid (GetControl ('Fliste'));
        {$ENDIF}
        if Grille  <>  NIL then Grille.OnDblClick  :=  GrilleDblClick;
        TFMul(Ecran).BInsert.OnClick := CreateForfait;
        //PT2 - Début
        // Dans le cas du calcul des frais par populations, on n'affiche plus l'établissement et on change de liste
        If VH_Paie.PGForGestFraisByPop Then
        Begin
               // Affichage des champs établissement
               SetControlVisible('PFF_POPULATION',True);
               SetControlVisible('TPFF_POPULATION',True);
               SetControlVisible('PFF_ETABLISSEMENT',False);
               SetControlVisible('TPFF_ETABLISSEMENT',False);
               // Mise à jour de la liste de la combo avec les populations actives
               SetControlProperty('PFF_POPULATION', 'Plus', ' AND PPC_PREDEFINI="'+GetPredefiniPopulation(TYP_POPUL_FORM_PREV)+'"');
               SetControlText('XX_WHERE','PFF_ETABLISSEMENT = "---"');

               // Changement de la liste
               TFMul(Ecran).SetDBListe('PGFORFAITFORMPOP');
        End
        Else
               SetControlText('XX_WHERE','PFF_ETABLISSEMENT <> "---"');
        //PT2 - Fin
end ;

procedure TOF_PGMULFORFAITFORM.GrilleDblClick(Sender : TObject);
var St : String;
begin
        {$IFDEF EAGLCLIENT}
        TFmul(Ecran).Q.TQ.Seek(TFmul(Ecran).FListe.Row-1) ;   //PT1
        {$ENDIF}
        If Q_Mul.FindField('PFF_LIEUFORM').AsString = '' then CreateForfait(Nil) //PT2
        Else
        begin
            //PT2 - Début
            If VH_Paie.PGForGestFraisByPop Then St := '---;'+Q_Mul.FindField('PFF_POPULATION').AsString
            Else St := Q_Mul.FindField('PFF_ETABLISSEMENT').AsString+';---';
            //PT2 - Fin
            St := St +';'+Q_Mul.FindField('PFF_LIEUFORM').AsString+';'+Q_Mul.FindField('PFF_MILLESIME').AsString;
            AGLLanceFiche ('PAY', 'FORFAITFORM', '',St,GetControlText('PFF_MILLESIME'));
        end;
        TFMul(Ecran).BChercheClick(TFMul(Ecran).BCherche);
end;

procedure TOF_PGMULFORFAITFORM.CreateForfait(Sender : TObject);
begin
        AGLLanceFiche ('PAY', 'FORFAITFORM','','', GetControlText('PFF_MILLESIME')+';ACTION=CREATION');
        TFMul(Ecran).BChercheClick(TFMul(Ecran).BCherche);
end;

Initialization
  registerclasses ( [ TOF_PGMULFORFAITFORM ] ) ;
end.
