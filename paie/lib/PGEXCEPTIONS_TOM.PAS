{***********UNITE*************************************************
Auteur  ...... :
Créé le ...... : 26/05/2005
Modifié le ... :   /  /
Description .. : Source TOM de la TABLE : PGEXCEPTIONS (PGEXCEPTIONS)
Mots clefs ... : TOM;PGEXCEPTIONS
*****************************************************************}
{
PT1   : 14/09/2005 PH V603 FQ 12515 En mono on autorise que la personnalisation
                           DOSSIER
PT2   : 29/09/2005 PH V603 FQ 12608 On coche rubrique imprimable si au moins 1
                           elt est imprimable
PT3   : 23/11/2005 PH V650 FQ 12723 En Mono, personnalisation des cotisations
                           impossibles
PT4   : 02/11/2006 GGS V_70 FQ 12608 Suite...
PT5   : 31/01/2007 GGS V_72 FQ 12694 Journal d'événements
PT6   : 08/03/2007 VG V_72 Ajout de la gestion des cumuls dans PGEXCEPTIONS
PT7   : 29/08/2007 VG V_80 Problème d'initialisation en mode PCL - FQ N°14707
PT8   : 06/11/2007 FC V_80 FQ 14908 Cumuls dans journal d'évènement
PT9   : 03/01/2008 PH V_80 On ne recharge pas les caractéristiques des rubriques en modification
                           de la personnalisation.
}

unit PGEXCEPTIONS_TOM;

interface

uses StdCtrls,
  Controls,
  Classes,
{$IFNDEF EAGLCLIENT}
  db,
  {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}
  Fiche,
{$ELSE}
  eFiche,
{$ENDIF}
  forms,
  sysutils,
  ComCtrls,
  HCtrls,
  HEnt1,
  HMsgBox,
  UTOM,
  UTob,
  PAIETOM;      //PT5

type
  TOM_PGEXCEPTIONS = class(PGTOM)     //PT5 PGTOM remplace TOM
    procedure OnNewRecord; override;
    procedure OnDeleteRecord; override;
    procedure OnUpdateRecord; override;
    procedure OnAfterUpdateRecord; override;
    procedure OnAfterDeleteRecord; override;
    procedure OnLoadRecord; override;
    procedure OnChangeField(F: TField); override;
    procedure OnArgument(S: string); override;
    procedure OnClose; override;
    procedure OnCancelRecord; override;
  private
    Nat: string;
    Trace: TStringList;            //PT5
    DerniereCreate: String;        //PT5
    LeStatut:TDataSetState; //PT8
  end;

implementation
uses pgoutils2, P5def;

procedure TOM_PGEXCEPTIONS.OnNewRecord;
begin
inherited;
if (Nat<>'') then
   SetField ('PEN_NATURERUB', Nat);
SetControlEnabled ('PEN_NATURERUB', FALSE);

if ((V_PGI.ModePCL<>'1') or (PGRendModeFonc()='MONO')) then // On préfixe en Dossier
   begin
   SetField ('PEN_PREDEFINI', 'DOS');
   SetControlEnabled ('PEN_PREDEFINI', FALSE);
{PT7
   if (Nat = 'REM') then
      SetControlProperty ('PEN_RUBRIQUE', 'Datatype', 'PGREMCEGID')
   else
   if (Nat = 'COT') then
      SetControlProperty ('PEN_RUBRIQUE', 'Datatype', 'PGCOTCEGID')
   else
   if (Nat = 'CUM') then
      begin
      SetField ('PEN_MODIFGESTASS', 'X');
      SetField ('PEN_PREDEFINI', 'STD');
      end;
}
   end;
//PT7
if (Nat='REM') then
   SetControlProperty ('PEN_RUBRIQUE', 'Datatype', 'PGREMCEGID')
else
if (Nat='COT') then
   SetControlProperty ('PEN_RUBRIQUE', 'Datatype', 'PGCOTCEGID')
//PT6
else
if (Nat='CUM') then
   begin
   SetField ('PEN_MODIFGESTASS', 'X');
   SetField ('PEN_PREDEFINI', 'STD');
   end;
//FIN PT6
//FIN PT7
end;

procedure TOM_PGEXCEPTIONS.OnDeleteRecord;
var
St,Code : String;
begin
inherited;
//PT5
{PT6
Trace.Add('SUPPRESSION PERSONNALISATION RUBRIQUE '+GetField('PEN_RUBRIQUE')+' '+ GetField('PEN_LIBELLE'));
CreeJnalEvt('003','086','OK',nil,nil,Trace);
Trace.free;
}
Trace := TStringList.Create;
if (Nat <> 'CUM') then
begin
   St:= 'RUBRIQUE ';
   Code := '086';   //PT8
end
else
begin
   St:= 'CUMUL ';
   Code := '096';   //PT8
end;

Trace.Add ('SUPPRESSION PERSONNALISATION '+St+GetField ('PEN_RUBRIQUE')+' '+
           GetField('PEN_LIBELLE'));
CreeJnalEvt('003', Code, 'OK', nil, nil, Trace);
FreeAndNil (Trace);
//FIN PT6
end;

procedure TOM_PGEXCEPTIONS.OnUpdateRecord;
var
St, NomTablette, Nature, Rubrique, Pred,Predefini: string;
Q : TQuery;
begin
inherited;
//PT8
if (DS.State in [dsInsert]) then
begin
  if GetField('PEN_PREDEFINI') = 'STD' then
    SetField('PEN_NODOSSIER', '000000')
  else
    SetField('PEN_NODOSSIER', PGRendNoDossier());
end;
//PT8
Nature:= GetField('PEN_NATURERUB');
Rubrique:= GetField('PEN_RUBRIQUE');
if (Nature <> 'CUM') then
   Pred:= GetField('PEN_PREDEFINI')
else
   Pred:= 'STD';
LeStatut := DS.State; //PT8
if (DS.State in [dsInsert]) then
   DerniereCreate:= GetField('PEN_RUBRIQUE');  //PT5
if PGRendModeFonc() = 'MONO' then
   begin // MONO personnalisation des rubriques CEGID uniquement et en mode Dossier
//PT6
   if (Nature<>'CUM') then
      begin
//FIN PT6
      if (Pred <> 'DOS') then
         begin
         LastErrorMsg:= 'Vous devez personnaliser la rubrique en dossier';
         LastError:= 1;
         end
      else
         begin
         if GetField('PEN_NATURERUB') = 'AAA' then
            NomTablette:= 'PGREMCEGID'
         else
            NomTablette:= 'PGCOTCEGID';
         St:= RechDom(NomTablette, GetField('PEN_RUBRIQUE'), FALSE);
         if ((St = '') or (St = 'Error')) then
            LastError:= 2
         else
            LastError:= 0;
         if (LastError <> 0) then
            LastErrorMsg := 'Vous ne pouvez personnaliser que des rubriques CEGID';
         end;
      end;   //PT6
   end
else
   begin
   if (GetField('PEN_NATURERUB')='AAA') then
      Q:= OpenSQl ('SELECT *'+
                   ' FROM REMUNERATION WHERE'+
                   ' ##PRM_PREDEFINI##'+
                   ' PRM_RUBRIQUE="'+GetField('PEN_RUBRIQUE')+'"', true)
   else
   if (GetField('PEN_NATURERUB')='COT') then                       //PT6
      Q:= OpenSQl ('SELECT *'+
                   ' FROM COTISATION WHERE'+
                   ' ##PCT_PREDEFINI## PCT_NATURERUB="COT" AND'+
                   ' PCT_RUBRIQUE="'+GetField('PEN_RUBRIQUE')+'"', true)
//PT6
   else
      Q:= OpenSQl ('SELECT *'+
                   ' FROM CUMULPAIE WHERE'+
                   ' ##PCL_PREDEFINI##'+
                   ' PCL_CUMULPAIE="'+GetField('PEN_RUBRIQUE')+'"', true);
//FIN PT6
   if not Q.EOF then
      begin
      if GetField('PEN_NATURERUB') = 'AAA' then
         Predefini:= Q.FindField ('PRM_PREDEFINI').AsString
      else
      if (GetField('PEN_NATURERUB')='COT') then            //PT6
         Predefini:= Q.FindField ('PCT_PREDEFINI').AsString
//PT6
      else
         Predefini:= Q.FindField ('PCL_PREDEFINI').AsString;
//FIN PT6
{PT6
      if (Predefini = 'STD') and (Pred <> 'DOS') then
}
      if ((Nature<>'CUM') and (Predefini='STD') and (Pred<>'DOS')) then
//FIN PT6
         begin
         LastError := 2;
         LastErrorMsg := 'Vous ne pouvez pas personnaliser en standard des rubriques standard';
         end;
      if (Predefini = 'DOS') then
         begin
         LastError := 2;
         LastErrorMsg := 'Vous ne pouvez pas personnaliser des rubriques dossier';
         end;
      end;
   FERME(Q);
   end;
end;

procedure TOM_PGEXCEPTIONS.OnAfterUpdateRecord;
var
even: boolean;      //PT5
  LaTable : String;
begin
inherited;
  Trace := TStringList.Create;    //PT6
  //DEB PT8
  if (Nat <> 'CUM') then
    LaTable := 'PGEXCEPTIONS'
  else
    LaTable := 'PGCUMEXCEPTIONS';
  even:= IsDifferent (dernierecreate,Latable,'PEN_RUBRIQUE',TFFiche(Ecran).LibelleName,Trace,TFFiche(Ecran),LeStatut);  //PT5  //PrefixeToTable(TFFiche(Ecran).TableName),
  //FIN PT8
{PT6
Trace.Free;
}
FreeAndNil (Trace);
//FIN PT6
end;

procedure TOM_PGEXCEPTIONS.OnAfterDeleteRecord;
begin
  inherited;
end;

procedure TOM_PGEXCEPTIONS.OnLoadRecord;
begin
  inherited;
  if not (DS.State in [dsInsert]) then DerniereCreate := '';      //PT5
end;

procedure TOM_PGEXCEPTIONS.OnChangeField(F: TField);
var
Q: TQuery;
Pred, Dossier: string;
begin
inherited;
if (F.FieldName='PEN_NATURERUB') then
   begin
   if GetField ('PEN_NATURERUB')='AAA' then
      begin
      SetControlVisible ('TBSHTREM', TRUE);
      SetControlVisible ('TBSHTCOT', FALSE);
      SetControlProperty ('PEN_RUBRIQUE', 'DataType', 'PGREMCEGSTD')
      end
   else
   if GetField ('PEN_NATURERUB')='COT' then           //PT6
      begin
      SetControlVisible ('TBSHTREM', FALSE);
      SetControlVisible ('TBSHTCOT', TRUE);
      SetControlProperty ('PEN_RUBRIQUE', 'DataType', 'PGCOTCEGSTD');
      end
//PT6
   else
   if (GetField ('PEN_NATURERUB')='CUM') then
      begin
      SetControlVisible ('TBSHTREM', FALSE);
      SetControlVisible ('TBSHTCOT', FALSE);
      end;
//FIN PT6
   end;

if ((F.FieldName='PEN_RUBRIQUE')) AND (DS.State = dsInsert) then // PT9
   begin
   if GetField ('PEN_NATURERUB')='AAA' then
      begin
      Q:= OpenSQl ('SELECT *'+
                   ' FROM REMUNERATION WHERE'+
                   ' ##PRM_PREDEFINI##'+
                   ' PRM_RUBRIQUE="'+GetField('PEN_RUBRIQUE')+'"', true);
      if not Q.EOF then
         begin
         Pred:= Q.FindField ('PRM_PREDEFINI').AsString;
         if Pred='CEG' then
            begin
            if (PGRendModeFonc ()='MONO') then
               Pred:= 'DOS'
            else
               Pred:= 'STD';
            Dossier:= '000000';
            end
         else
            begin
            Pred:= 'DOS';
            Dossier:= PGRendNoDossier();
            end;
         SetField ('PEN_PREDEFINI', Pred);
         SetField ('PEN_NODOSSIER', Dossier);
         SetField ('PEN_LIBELLE', Q.FindField ('PRM_LIBELLE').AsString);
         SetField ('PEN_IMPRIMABLE', Q.FindField ('PRM_IMPRIMABLE').AsString);
         SetField ('PEN_BASEIMPRIMABLE',
                   Q.FindField ('PRM_BASEIMPRIMABLE').AsString);
         SetField ('PEN_DECBASE', Q.FindField ('PRM_DECBASE').AsString);
         SetField ('PEN_TAUXIMPRIMABLE',
                   Q.FindField ('PRM_TAUXIMPRIMABLE').AsString);
         SetField ('PEN_COEFFIMPRIM', Q.FindField ('PRM_COEFFIMPRIM').AsString);
         SetField ('PEN_DECTAUX', Q.FindField ('PRM_DECTAUX').AsInteger);
         SetField ('PEN_DECCOEFF', Q.FindField ('PRM_DECCOEFF').AsInteger);
         end
      else
         SetFocusControl ('PEN_RUBRIQUE');
      FERME(Q);
      end
   else
   if (GetField ('PEN_NATURERUB')='COT') then  //PT6
      begin
      Q:= OpenSQl ('SELECT *'+
                   ' FROM COTISATION WHERE'+
                   ' ##PCT_PREDEFINI## PCT_NATURERUB="COT" AND'+
                   ' PCT_RUBRIQUE="'+GetField ('PEN_RUBRIQUE')+'"', true);
      if not Q.EOF then
         begin
         Pred:= Q.FindField('PCT_PREDEFINI').AsString;
         if Pred = 'CEG' then
            begin
            if (PGRendModeFonc ()='MONO') then
               Pred:= 'DOS'
            else
               Pred:= 'STD';
            Dossier:= '000000';
            end
         else
            begin
            Pred:= 'DOS';
            Dossier:= PGRendNoDossier();
            end;
         SetField ('PEN_PREDEFINI', Pred);
         SetField ('PEN_NODOSSIER', Dossier);
         SetField ('PEN_LIBELLE', Q.FindField ('PCT_LIBELLE').AsString);
         SetField ('PEN_IMPRIMABLE', Q.FindField ('PCT_IMPRIMABLE').AsString);
         SetField ('PEN_BASEIMPRIMABLE', Q.FindField ('PCT_BASEIMP').AsString);
         SetField ('PEN_DECBASE', Q.FindField ('PCT_DECBASE').AsString);
         SetField ('PEN_TXSALIMP', Q.FindField ('PCT_TXSALIMP').AsString);
         SetField ('PEN_FFSALIMP', Q.FindField ('PCT_FFSALIMP').AsString);
         SetField ('PEN_TXPATIMP', Q.FindField ('PCT_TXPATIMP').AsString);
         SetField ('PEN_FFPATIMP', Q.FindField ('PCT_FFPATIMP').AsString);
         SetField ('PEN_ORGANISME', Q.FindField ('PCT_ORGANISME').AsString);
         SetField ('PEN_DECTXSAL', Q.FindField ('PCT_DECTXSAL').AsInteger);
         SetField ('PEN_DECMTSAL', Q.FindField ('PCT_DECMTSAL').AsInteger);
         SetField ('PEN_DECTXPAT', Q.FindField ('PCT_DECTXPAT').AsInteger);
         SetField ('PEN_DECMTPAT', Q.FindField ('PCT_DECMTPAT').AsInteger);
         end
      else
         SetFocusControl('PEN_RUBRIQUE');
      FERME(Q);
      end
//PT6
   else
   if (GetField ('PEN_NATURERUB')='CUM') then
      begin
      Q:= OpenSQl ('SELECT *'+
                   ' FROM CUMULPAIE WHERE'+
                   ' ##PCL_PREDEFINI##'+
                   ' PCL_CUMULPAIE="'+GetField ('PEN_RUBRIQUE')+'"', true);
      if not Q.EOF then
         begin
         Dossier:= PGRendNoDossier();
         SetField ('PEN_PREDEFINI', 'STD');
         SetField ('PEN_NODOSSIER', Dossier);
         SetField ('PEN_LIBELLE', Q.FindField ('PCL_LIBELLE').AsString);
         end
      else
         SetFocusControl('PEN_RUBRIQUE');
      FERME(Q);
      end;
//FIN PT6
   end;

if (F.FieldName='PEN_PREDEFINI') then
   begin
   if GetField('PEN_PREDEFINI')='STD' then
      SetField ('PEN_NODOSSIER', '000000')
   else
      SetField ('PEN_NODOSSIER', PGRendNoDossier());
   end;

//PT4  if (POS('IMP', F.FieldName) > 0) and (F.FieldName <> 'PEN_IMPRIMABLE') then setField('PEN_IMPRIMABLE', 'X');
if (POS ('IMP', F.FieldName)>0) and (F.FieldName<>'PEN_IMPRIMABLE') then
   begin
   if GetField (F.FieldName)='X' then
      SetField ('PEN_IMPRIMABLE', 'X'); //PT4
   end;

if (F.FieldName='PEN_IMPRIMABLE') then
   begin
   if (GetField ('PEN_IMPRIMABLE')='-') then
      begin
      SetField ('PEN_TXSALIMP', '-');
      SetField ('PEN_FFSALIMP', '-');
      SetField ('PEN_TXPATIMP', '-');
      SetField ('PEN_FFPATIMP', '-');
      SetField ('PEN_BASEIMPRIMABLE', '-');
      SetField ('PEN_TAUXIMPRIMABLE', '-');
      SetField ('PEN_COEFFIMPRIM', '-');
      end;
   end;
end;

procedure TOM_PGEXCEPTIONS.OnArgument(S: string);
begin
inherited;
Nat:= ReadTokenSt (S); // Action
Nat:= ReadTokenSt (S); // Nature de la rubrique
SetControlEnabled ('PEN_NATURERUB', FALSE);
{PT6
Trace:= TStringList.Create ;         //PT5
}
end;

procedure TOM_PGEXCEPTIONS.OnClose;
begin
  inherited;
end;

procedure TOM_PGEXCEPTIONS.OnCancelRecord;
begin
  inherited;
end;

initialization
  registerclasses([TOM_PGEXCEPTIONS]);
end.

