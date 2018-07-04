{***********UNITE*************************************************
Auteur  ...... : 
Créé le ...... : 11/12/2002
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : PGANALYSEFORMATION ()
Mots clefs ... : TOF;PGANALYSEFORMATION
*****************************************************************
PT1 04/11/2003 JL V_50 Ajout table SESSIONSTAGE dans la requête
PT2 27/11/2003 JL V_50 Ajout affichage champs travail salariés
---- JL 17/10/2006 Modification contrôle des exercices de formations -----
}
Unit UTofPG_AnalyseFormation ;

Interface

Uses Classes,
{$IFNDEF EAGLCLIENT}
     db,{$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}
{$ELSE}
     UTob,
{$ENDIF}
     sysutils,HCtrls,UTOF,PGOutilsFormation,EntPaie,Cube,P5DEF ;

Type
  TOF_PGANALYSEFORMATION = Class (TOF)
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
  end ;

Implementation

procedure TOF_PGANALYSEFORMATION.OnLoad;
var Where : String;
begin
        Inherited ;
        If GetControl('DECLARATION')<> Nil then
        begin
                If GetControlText('DECLARATION') = 'TOUS' then Where := ''
                else If GetControlText('DECLARATION') = 'INCLUS' then Where := 'PSS_INCLUSDECL="X"'
                else Where := 'PSS_INCLUSDECL="-"';
                setControlText('XX_WHERE',Where);
        end;
end;

procedure TOF_PGANALYSEFORMATION.OnArgument (S : String ) ;
var Num : Integer;
    DateDeb,DateFin : TDateTime;
begin
  Inherited ;
        For Num := 1 to VH_Paie.NBFormationLibre do
        begin
                if Num > 8 then Break;
                VisibiliteChampFormation (IntToStr(Num),GetControl ('PFO_FORMATION'+IntToStr(Num)),GetControl ('TPFO_FORMATION'+IntToStr(Num)));
        end;
        DateDeb := Date;
        DateFin := Date;
        RendMillesimeRealise(dateDeb,DateFin);
        SetControlText('PFO_DATEDEBUT',DateToStr(DateDeb));
        SetControlText('PFO_DATEDEBUT_',DateToStr(DateFin));
        TFCube(Ecran).FromSQL := 'FORMATIONS LEFT JOIN SESSIONSTAGE ON PFO_CODESTAGE=PSS_CODESTAGE AND PFO_ORDRE=PSS_ORDRE AND PFO_MILLESIME=PSS_MILLESIME LEFT JOIN SALARIES ON PFO_SALARIE=PSA_SALARIE'; //PT1
        if GetControl('PFO_TRAVAILN1') <> nil then              //PT2
        begin
                For Num  :=  1 to VH_Paie.PGNbreStatOrg do
                begin
                        if Num >4 then Break;
                        VisibiliteChampSalarie (IntToStr(Num),GetControl ('PFO_TRAVAILN'+IntToStr(Num)),GetControl ('TPFO_TRAVAILN'+IntToStr(Num)));
                end;
                VisibiliteStat (GetControl ('PFO_CODESTAT'),GetControl ('TPFO_CODESTAT')) ;
        end;
        If GetControl('DECLARATION')<> Nil then SetControltext('DECLARATION','TOUS');
end ;

Initialization
  registerclasses ( [ TOF_PGANALYSEFORMATION ] ) ;
end.
