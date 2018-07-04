{***********UNITE*************************************************
Auteur  ...... : JL
Créé le ...... : 26/09/2002
Modifié le ... :   /  /
Description .. : TOT gestion des tablettes libres
Mots clefs ... : PAIE
*****************************************************************
PT1 22/12/2004 V_60 JL Ajout gestion tablette libres scoring
}
unit UTOTPGFormationTabletteLibre;

Interface

Uses Classes,sysutils,
{$IFNDEF EAGLCLIENT}
     db,{$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}
{$ELSE}
     UTOB,eTablette,
{$ENDIF}
     HCtrls,HMsgBox,UTOT,Entpaie;

Type
  TOT_PGFORMATION1 = Class ( TOT )
    function GetTitre : HString; override ;
    procedure OnClose; override ;
    Procedure OnDeleteRecord ; override ;
  end ;

  TOT_PGFORMATION2 = Class ( TOT )
    function GetTitre : HString; override ;
    procedure OnClose; override ;
    Procedure OnDeleteRecord ; override ;
  end ;

  TOT_PGFORMATION3 = Class ( TOT )
    function GetTitre : HString; override ;
    procedure OnClose; override ;
    Procedure OnDeleteRecord ; override ;
  end ;

    TOT_PGFORMATION4 = Class ( TOT )
   function GetTitre : HString; override ;
   procedure OnClose; override ;
   Procedure OnDeleteRecord ; override ;
  end ;

    TOT_PGFORMATION5 = Class ( TOT )
   function GetTitre : HString; override ;
   procedure OnClose; override ;
   Procedure OnDeleteRecord ; override ;
  end ;

    TOT_PGFORMATION6 = Class ( TOT )
   function GetTitre : HString; override ;
   procedure OnClose; override ;
   Procedure OnDeleteRecord ; override ;
  end ;

    TOT_PGFORMATION7 = Class ( TOT )
   function GetTitre : HString; override ;
   procedure OnClose; override ;
   Procedure OnDeleteRecord ; override ;
  end ;

    TOT_PGFORMATION8 = Class ( TOT )
   function GetTitre : HString; override ;
   procedure OnClose; override ;
   Procedure OnDeleteRecord ; override ;
  end ;

  TOT_PGFRAISSALFORM = Class ( TOT )
   Procedure OnUpdateRecord ; override ;
   Procedure OnDeleteRecord ; override ;
  end ;

  TOT_PGMOTIFETATINSC1 = Class ( TOT )
   procedure OnNewRecord ; override ;
  end;

    TOT_PGMOTIFETATINSC2 = Class ( TOT )
   procedure OnNewRecord ; override ;
  end;

    TOT_PGMOTIFETATINSC3 = Class ( TOT )
   procedure OnNewRecord ; override ;
  end;

  {SCORING}
  TOT_PGFSCORING1 = Class ( TOT )
   procedure OnNewRecord ; override ;
  end;

  TOT_PGFSCORING2 = Class ( TOT )
   procedure OnNewRecord ; override ;
  end;

  TOT_PGFSCORING3 = Class ( TOT )
   procedure OnNewRecord ; override ;
  end;

  TOT_PGFSCORING4 = Class ( TOT )
   procedure OnNewRecord ; override ;
  end;

  TOT_PGFSCORING5 = Class ( TOT )
   procedure OnNewRecord ; override ;
  end;

implementation
{ TOT_PGFORMATION1 }

function TOT_PGFORMATION1.GetTitre: HString;
begin
        result := VH_Paie.FormationLibre1;
end;

procedure TOT_PGFORMATION1.OnClose;
begin
Inherited ;
        AvertirTable('PGFORMATION1');
end;

Procedure TOT_PGFORMATION1.OnDeleteRecord ;
begin
Inherited ;
        If ExisteSQL('SELECT PST_CODESTAGE FROM STAGE WHERE PST_FORMATION1="'+GetField('CC_CODE')+'"') then
        begin
                PGIBox('Suppression impossible car il existe des formations associées à cette tablette libre','Saisie '+ VH_Paie.FormationLibre1);
                LastError := 1;
                Exit;
        end;
end;

{ TOT_PGFORMATION2 }
function TOT_PGFORMATION2.GetTitre: HString;
begin
        result := VH_Paie.FormationLibre2;
end;

procedure TOT_PGFORMATION2.OnClose;
begin
Inherited ;
        AvertirTable('PGFORMATION2');
end;

Procedure TOT_PGFORMATION2.OnDeleteRecord ;
begin
Inherited ;
        If ExisteSQL('SELECT PST_CODESTAGE FROM STAGE WHERE PST_FORMATION2="'+GetField('CC_CODE')+'"') then
        begin
                PGIBox('Suppression impossible car il existe des formations associées à cette tablette libre','Saisie '+ VH_Paie.FormationLibre2);
                LastError := 1;
                Exit;
        end;
end;

{ TOT_PGFORMATION3 }
function TOT_PGFORMATION3.GetTitre: HString;
begin
        result := VH_Paie.FormationLibre3;
end;

procedure TOT_PGFORMATION3.OnClose;
begin
Inherited ;
        AvertirTable('PGFORMATION3');
end;

Procedure TOT_PGFORMATION3.OnDeleteRecord ;
begin
Inherited ;
        If ExisteSQL('SELECT PST_CODESTAGE FROM STAGE WHERE PST_FORMATION3="'+GetField('CC_CODE')+'"') then
        begin
                PGIBox('Suppression impossible car il existe des formations associées à cette tablette libre','Saisie '+ VH_Paie.FormationLibre3);
                LastError := 1;
                Exit;
        end;
end;

{ TOT_PGFORMATION4 }
function TOT_PGFORMATION4.GetTitre: HString;
begin
        result := VH_Paie.FormationLibre4;
end;

procedure TOT_PGFORMATION4.OnClose;
begin
Inherited ;
        AvertirTable('PGFORMATION4');
end;

Procedure TOT_PGFORMATION4.OnDeleteRecord ;
begin
Inherited ;
        If ExisteSQL('SELECT PST_CODESTAGE FROM STAGE WHERE PST_FORMATION4="'+GetField('CC_CODE')+'"') then
        begin
                PGIBox('Suppression impossible car il existe des formations associées à cette tablette libre','Saisie '+ VH_Paie.FormationLibre4);
                LastError := 1;
                Exit;
        end;
end;

{ TOT_PGFORMATION5 }
function TOT_PGFORMATION5.GetTitre: HString;
begin
        result := VH_Paie.FormationLibre5;
end;

procedure TOT_PGFORMATION5.OnClose;
begin
Inherited ;
        AvertirTable('PGFORMATION5');
end;

Procedure TOT_PGFORMATION5.OnDeleteRecord ;
begin
Inherited ;
        If ExisteSQL('SELECT PST_CODESTAGE FROM STAGE WHERE PST_FORMATION5="'+GetField('CC_CODE')+'"') then
        begin
                PGIBox('Suppression impossible car il existe des formations associées à cette tablette libre','Saisie '+ VH_Paie.FormationLibre5);
                LastError := 1;
                Exit;
        end;
end;

{ TOT_PGFORMATION6 }
function TOT_PGFORMATION6.GetTitre: HString;
begin
        result := VH_Paie.FormationLibre6;
end;

procedure TOT_PGFORMATION6.OnClose;
begin
Inherited ;
        AvertirTable('PGFORMATION6');
end;

Procedure TOT_PGFORMATION6.OnDeleteRecord ;
begin
Inherited ;
        If ExisteSQL('SELECT PST_CODESTAGE FROM STAGE WHERE PST_FORMATION6="'+GetField('CC_CODE')+'"') then
        begin
                PGIBox('Suppression impossible car il existe des formations associées à cette tablette libre','Saisie '+ VH_Paie.FormationLibre6);
                LastError := 1;
                Exit;
        end;
end;

{ TOT_PGFORMATION7 }
function TOT_PGFORMATION7.GetTitre: HString;
begin
        result := VH_Paie.FormationLibre7;
end;

procedure TOT_PGFORMATION7.OnClose;
begin
Inherited ;
        AvertirTable('PGFORMATION7');
end;

Procedure TOT_PGFORMATION7.OnDeleteRecord ;
begin
Inherited ;
        If ExisteSQL('SELECT PST_CODESTAGE FROM STAGE WHERE PST_FORMATION7="'+GetField('CC_CODE')+'"') then
        begin
                PGIBox('Suppression impossible car il existe des formations associées à cette tablette libre','Saisie '+ VH_Paie.FormationLibre7);
                LastError := 1;
                Exit;
        end;
end;

{ TOT_PGFORMATION8 }
function TOT_PGFORMATION8.GetTitre: HString;
begin
        result := VH_Paie.FormationLibre8;
end;

procedure TOT_PGFORMATION8.OnClose;
begin
Inherited ;
        AvertirTable('PGFORMATION8');
end;

Procedure TOT_PGFORMATION8.OnDeleteRecord ;
begin
Inherited ;
        If ExisteSQL('SELECT PST_CODESTAGE FROM STAGE WHERE PST_FORMATION8="'+GetField('CC_CODE')+'"') then
        begin
                PGIBox('Suppression impossible car il existe des formations associées à cette tablette libre','Saisie '+ VH_Paie.FormationLibre8);
                LastError := 1;
                Exit;
        end;
end;

{TOT_PGFRAISSALFORM}
procedure TOT_PGFRAISSALFORM.OnUpdateRecord;
var NbFrais:Integer;
    Q:TQuery;
begin
Inherited ;
        NbFrais := 0;
        Q := OpenSQL('SELECT COUNT(CC_CODE) AS NBFRAIS FROM CHOIXCOD WHERE CC_TYPE="PFA"',True);
        If Not Q.Eof then NbFrais := Q.FindField('NBFRAIS').AsInteger;
        Ferme(Q);
        If NbFrais>=15 then
        begin
                PGIBox('Impossible de créer un nouveau frais, #13#10 vous avez atteint le maximum de frais possibles : 15','Saisie des frais');
                LastError := 1;
                Exit;
        end
        Else
        begin
                If NbFrais>VH_Paie.PGFNbFraisLibre then PGIBox('Attention, seulement '+IntToStr(VH_Paie.PGFNbFraisLibre)+' sont gérés dans les paramètres société#13#10 vous devrez modifier le nombre pour pouvoir afficher le nouveau frais','Saisie des frais');
        end;
end;

Procedure TOT_PGFRAISSALFORM.OnDeleteRecord ;
begin
Inherited ;
        If ExisteSQL('SELECT PFS_FRAISSALFOR FROM FRAISSALFORM WHERE PFS_FRAISSALFOR="'+GetField('CC_CODE')+'"') then
        begin
                PGIBox('Suppression impossible car il existe des enregistrements pour ce type de frais.','Saisie des libellés des frais de déplacement');
                LastError := 1;
                Exit;
        end;
end;

{TOT_PGMOTIFETATINSC1}
procedure TOT_PGMOTIFETATINSC1.OnNewRecord ;
begin
Inherited ;
        SetField('CC_LIBRE','VAL');
end;

{TOT_PGMOTIFETATINSC2}
procedure TOT_PGMOTIFETATINSC2.OnNewRecord ;
begin
Inherited ;
        SetField('CC_LIBRE','REP');
end;

{TOT_PGMOTIFETATINSC3}
procedure TOT_PGMOTIFETATINSC3.OnNewRecord ;
begin
Inherited ;
        SetField('CC_LIBRE','REF');
end;

{SCORING}
{TOT_PGFSCORING1}
procedure TOT_PGFSCORING1.OnNewRecord ;
begin
Inherited ;
        SetField('CC_LIBRE','SC1');
end;

{TOT_PGFSCORING2}
procedure TOT_PGFSCORING2.OnNewRecord ;
begin
Inherited ;
        SetField('CC_LIBRE','SC2');
end;

{TOT_PGFSCORING3}
procedure TOT_PGFSCORING3.OnNewRecord ;
begin
Inherited ;
        SetField('CC_LIBRE','SC3');
end;

{TOT_PGFSCORING4}
procedure TOT_PGFSCORING4.OnNewRecord ;
begin
Inherited ;
        SetField('CC_LIBRE','SC4');
end;

{TOT_PGFSCORING5}
procedure TOT_PGFSCORING5.OnNewRecord ;
begin
Inherited ;
        SetField('CC_LIBRE','SC5');
end;

Initialization
  registerclasses ( [ TOT_PGFORMATION1,TOT_PGFORMATION2,TOT_PGFORMATION3,TOT_PGFORMATION4,TOT_PGFORMATION5,TOT_PGFORMATION6,TOT_PGFORMATION7,TOT_PGFORMATION8,TOT_PGFRAISSALFORM,TOT_PGMOTIFETATINSC1,TOT_PGMOTIFETATINSC2,TOT_PGMOTIFETATINSC3,TOT_PGFSCORING1,TOT_PGFSCORING2,TOT_PGFSCORING3,TOT_PGFSCORING4,TOT_PGFSCORING5] ) ;
end.

