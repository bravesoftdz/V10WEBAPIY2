{***********UNITE*************************************************
Auteur  ...... : JL
Créé le ...... : 30/09/2003
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : CUDE_PREVFORM ()
Mots clefs ... : TOF;PGANALYSEPREVISIONNELFORM
*****************************************************************
---- JL 17/10/2006 Modification contrôle des exercices de formations -----
}
Unit UTofPG_AnalysePrevisonnelForm ;

Interface

Uses Classes,
{$IFNDEF EAGLCLIENT}
     db,{$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}
{$ELSE}
     UTob,
{$ENDIF}
     sysutils,HCtrls,UTOF,PGOutilsFormation,EntPaie ;

Type
  TOF_PGANALYSEPREVFORM = Class (TOF)
    procedure OnArgument (S : String ) ; override ;
  end ;

Implementation

procedure TOF_PGANALYSEPREVFORM.OnArgument (S : String ) ;
var Num : Integer;
begin
  Inherited ;
        For Num := 1 to VH_Paie.NBFormationLibre do
        begin
                if Num > 8 then Break;
                VisibiliteChampFormation (IntToStr(Num),GetControl ('PFI_FORMATION'+IntToStr(Num)),GetControl ('TPFI_FORMATION'+IntToStr(Num)));
        end;
        SetControltext('PFI_MILLESIME',RendMillesimePrevisionnel);
end ;

Initialization
  registerclasses ( [ TOF_PGANALYSEPREVFORM ] ) ;
end.
