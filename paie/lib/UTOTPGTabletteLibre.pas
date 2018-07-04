{***********UNITE*************************************************
Auteur  ...... : PH
Créé le ...... : 12/09/2001
Modifié le ... :   /  /
Description .. : TOT gestion des tablettes libres
Mots clefs ... : PAIE
*****************************************************************}
{
PT1   : 06/05/2004 PH V_50 Traitement Tablette catégorie du bilan social
PT2   : 13/10/2006 SB V_70 FQ 13560 limitation du nb de catégorie
PT3   : 31/10/2006 SB V_70 Reprise FQ 13560
PT4   : 22/07/2008 VG V_80 Suppression de la limite à 10 Catégories Bilan social
                           FQ N°15387
}
unit UTOTPGTabletteLibre;

interface

uses StdCtrls, Controls, Classes, forms, sysutils, ComCtrls,
  {$IFNDEF EAGLCLIENT}
  db, {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF} HDB, DBCtrls, Tablette,
  {$ELSE}
  UTOB, eTablette,
  {$ENDIF}
  HCtrls, HEnt1, HMsgBox, UTOT, HQry, Entpaie;

//  PT1 06/05/04 V5_00  PH Traitement Tablette catégorie du bilan social
type
  TOT_PGCATBILAN = class(TOT)
    procedure OnNewRecord; override;
    procedure OnDeleteRecord ; override;
    procedure OnUpdateRecord; override; { PT3 }
    private
    NbreCat: Integer;
  end;
// FIN PT1
type
  TOT_PGLIBREPCMB1 = class(TOT)
    function GetTitre: Hstring; override;
  end;

  TOT_PGLIBREPCMB2 = class(TOT)
    function GetTitre: Hstring; override;
  end;

  TOT_PGLIBREPCMB3 = class(TOT)
    function GetTitre: Hstring; override;
  end;

  TOT_PGLIBREPCMB4 = class(TOT)
    function GetTitre: Hstring; override;
  end;

  TOT_PGTRAVAILN1 = class(TOT)
    function GetTitre: Hstring; override;
  end;
  TOT_PGTRAVAILN2 = class(TOT)
    function GetTitre: Hstring; override;
  end;
  TOT_PGTRAVAILN3 = class(TOT)
    function GetTitre: Hstring; override;
  end;
  TOT_PGTRAVAILN4 = class(TOT)
    function GetTitre: Hstring; override;
  end;

implementation
{ TOT_PGLIBREPCMB1 }

function TOT_PGLIBREPCMB1.GetTitre: Hstring;
begin
  result := VH_Paie.PgLibCombo1;
  //RechDom('PGLIBREPCMB1',VH_Paie.PgLibCombo1,False);
end;



{ TOT_PGLIBREPCMB2 }

function TOT_PGLIBREPCMB2.GetTitre: Hstring;
begin
  result := VH_Paie.PgLibCombo2;
end;

{ TOT_PGLIBREPCMB3 }

function TOT_PGLIBREPCMB3.GetTitre: Hstring;
begin
  result := VH_Paie.PgLibCombo3;
end;

{ TOT_PGLIBREPCMB4 }

function TOT_PGLIBREPCMB4.GetTitre: Hstring;
begin
  result := VH_Paie.PgLibCombo4;
end;


{ TOT_PGTRAVAILN1 }

function TOT_PGTRAVAILN1.GetTitre: Hstring;
begin
  result := VH_Paie.PGLibelleOrgStat1;
end;

{ TOT_PGTRAVAILN2 }

function TOT_PGTRAVAILN2.GetTitre: Hstring;
begin
  result := VH_Paie.PGLibelleOrgStat2;
end;

{ TOT_PGTRAVAILN3 }

function TOT_PGTRAVAILN3.GetTitre: Hstring;
begin
  result := VH_Paie.PGLibelleOrgStat3;
end;

{ TOT_PGTRAVAILN4 }

function TOT_PGTRAVAILN4.GetTitre: Hstring;
begin
  result := VH_Paie.PGLibelleOrgStat4;
end;

{ TOT_PGTYPEORGANISME }
procedure TOT_PGCATBILAN.OnDeleteRecord;
begin
  inherited;
  if GetField('CC_CODE')='000' then
    Begin
    LastError := 1 ;
    PgiBox('Suppression impossible. L''enregistrement "000" est prédéfini CEGID.','Catégories');
    Exit;
    End;

  If ExisteSql('SELECT * FROM SALARIES WHERE PSA_CATBILAN="'+GetField('CC_CODE')+'"') then
    Begin
    LastError := 1 ;
    PgiBox('Suppression impossible. Des salariés sont affectés à cette catégorie bilan social.','Catégories');
    End;

end;

//Traitement Tablette catégorie du bilan social
procedure TOT_PGCATBILAN.OnNewRecord;
begin
inherited;
{PT4
  Q := OpenSQL ('SELECT COUNT (*) NBRE FROM CHOIXCOD WHERE CC_TYPE="PCL"', True) ;
  if NOT Q.EOF then NbreCat := Q.FindField ('NBRE').AsInteger
  else NbreCat := 0 ;
  FERME (Q);

  if NbreCat >= 10 then
  begin
    PGiBox ('Vous ne pouvez pas créer plus de 10 catégories.', 'Catégories du bilan social');
    LastError := 1 ;
  end ;
}  
end;

procedure TOT_PGCATBILAN.OnUpdateRecord;
begin
inherited;
{PT4
  if NbreCat >= 10 then
  begin
    PGiBox ('Vous ne pouvez pas créer plus de 10 catégories.', 'Catégories du bilan social');
    LastError := 1 ;
  end ;
}  
end;


initialization
  registerclasses([TOT_PGLIBREPCMB1, TOT_PGLIBREPCMB2, TOT_PGLIBREPCMB3, TOT_PGLIBREPCMB4, TOT_PGTRAVAILN1, TOT_PGTRAVAILN2, TOT_PGTRAVAILN3, TOT_PGTRAVAILN4, TOT_PGCATBILAN]);
end.

