{***********UNITE*************************************************
Auteur  ...... : JL
Créé le ...... : 30/09/2003
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : CUDE_ANIMAT ()
Mots clefs ... : TOF;PGANALYSEANIMATEURS
*****************************************************************
PT1 27/11/2003 JL V_50 Ajout table SESSIONSTAGE dans la requête
---- JL 17/10/2006 Modification contrôle des exercices de formations -----
}
Unit UTofPG_AnalyseAnimateurs ;

Interface

Uses Classes,
{$IFNDEF EAGLCLIENT}
     db,{$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}
{$ELSE}
     UTob,
{$ENDIF}
     sysutils,HCtrls,HEnt1,UTOF,P5DEF,EntPaie,Cube,PGOutilsFormation ;

Type
  TOF_PGANALYSEANIMAT = Class (TOF)
    procedure OnArgument (S : String ) ; override ;
  end ;

Implementation

procedure TOF_PGANALYSEANIMAT.OnArgument (S : String ) ;
var Num : Integer;
    DateDeb,DateFin : TDateTime;
begin
  Inherited ;
        For Num  :=  1 to VH_Paie.PGNbreStatOrg do
        begin
                if Num >4 then Break;
                VisibiliteChampSalarie (IntToStr(Num),GetControl ('PSA_TRAVAILN'+IntToStr(Num)),GetControl ('TPSA_TRAVAILN'+IntToStr(Num)));
        end;
        DateDeb := V_PGI.DateEntree;
        DateFin := V_PGI.DateEntree;
        RendMillesimeRealise(dateDeb,DateFin);
        SetControlText('PAN_DATEDEBUT',DateToStr(DateDeb));
        SetControlText('PAN_DATEDEBUT_',DateToStr(DateFin));
        TFCube(Ecran).FromSQL := 'SESSIONANIMAT LEFT JOIN SESSIONSTAGE ON PAN_CODESTAGE=PSS_CODESTAGE AND PAN_ORDRE=PSS_ORDRE AND PAN_MILLESIME=PSS_MILLESIME LEFT JOIN SALARIES ON PAN_SALARIE=PSA_SALARIE'; //PT1
end ;

Initialization
  registerclasses ( [ TOF_PGANALYSEANIMAT ] ) ;
end.
