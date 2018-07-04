unit PgPlanningUnifie;

interface
uses
  SysUtils,
  HCtrls,
{$IFNDEF EAGLCLIENT}
  {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}
{$ENDIF}
  Utob,
  YPlanning,
  YRessource;

Type TypePGYPL = record
     GUID : string;             //Guid
     Salarie : string;          //Code salarié
     DateDebutAbs : TDateTime;  //Date de début d'absence
     DateFinAbs : TDateTime;    //Date de fin d'absence
     Libelle : string;          //Libellé de l'absence
end;

  
function CreatePGYPL (Tob_CP : Tob; PGYPL : TypePGYPL): integer;     { Création dans YPLANNING }
function DeletePGYPL (Const GuidOri, GuidYRS : string): integer;     { Suppression dans YPLANNING }


implementation
{***********A.G.L.Privé.*****************************************
Auteur  ...... : Vincent GALLIOT
Créé le ...... : 15/10/2007
Modifié le ... :   /  /
Description .. : Insertion d'un enregistrement dans le planning unifié dans la 
Suite ........ : table YPLANNING
Mots clefs ... : PGPLANNING
*****************************************************************}
function CreatePGYPL (Tob_CP : Tob; PGYPL : TypePGYPL): integer;
var
TobYPL : Tob;
Guid, Libelle, Salarie : string;
DateDebutAbs, DateFinAbs : TDateTime;
begin
result:= 0;
if not GestionYPlanning then
   exit;

if (Tob_CP<>nil) then
   begin
   Guid:= Tob_CP.GetValue ('PCN_GUID');
   Salarie:= Tob_CP.GetValue ('PCN_SALARIE');
   DateDebutAbs:= Tob_CP.GetValue ('PCN_DATEDEBUTABS');
   DateFinAbs:= Tob_CP.GetValue ('PCN_DATEFINABS');
   Libelle:= Tob_CP.GetValue ('PCN_LIBELLE');
   end
else
   begin
   Guid:= PGYPL.GUID;
   Salarie:= PGYPL.Salarie;
   DateDebutAbs:= PGYPL.DateDebutAbs;
   DateFinAbs:= PGYPL.DateFinAbs;
   Libelle:= PGYPL.Libelle;
   end;

TobYPL:= TOB.Create ('YPLANNING', nil, -1);
try
   LoadTobYPL ('YPL_PREFIXE="'+cStTablePaie+'" AND'+
               ' YPL_GUIDORI="'+Guid+'"', TobYPL);
   if (TobYPL.Detail.Count<1) then
      begin
      with TOB.Create ('#YPLANNING', TobYPL, -1) do
           begin
           AddChampSupValeur ('YPL_PREFIXE', cStTablePaie);
           AddChampSupValeur ('YPL_GUIDORI', Guid);
           AddChampSupValeur ('YPL_GUIDYRS', GetYRS_GUID (Salarie, '', ''));
           AddChampSupValeur ('YPL_DATEDEBUT', DateDebutAbs);
           AddChampSupValeur ('YPL_DATEFIN', DateFinAbs);
           AddChampSupValeur ('YPL_TYPEYPL', '');
           AddChampSupValeur ('YPL_STATUTYPL', '');
           AddChampSupValeur ('YPL_LIBELLE', Libelle);
           AddChampSupValeur ('YPL_ABREGE',  Libelle);
           AddChampSupValeur ('YPL_PRIVE',   '-');
           end;
      result:= CreateYPL (TobYPL.Detail[0]);
      end
   else
      begin
      with TobYPL.Detail[0] do
           begin
           PutValue ('YPL_PREFIXE', cStTablePaie);
           PutValue ('YPL_GUIDORI', Guid);
           PutValue ('YPL_GUIDYRS', GetYRS_GUID (Salarie, '', ''));
           PutValue ('YPL_DATEDEBUT', DateDebutAbs);
           PutValue ('YPL_DATEFIN', DateFinAbs);
           PutValue ('YPL_TYPEYPL', '');
           PutValue ('YPL_STATUTYPL', '');
           PutValue ('YPL_LIBELLE', Libelle);
           PutValue ('YPL_ABREGE', Libelle);
           PutValue ('YPL_PRIVE', '-');
           end;
      result:= UpdateYPL (tobYPL.Detail[0]);
      end;
finally
   FreeAndNil (TobYPL);
   end;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : Vincent GALLIOT
Créé le ...... : 15/10/2007
Modifié le ... :   /  /
Description .. : Suppression d'un enregistrement dans le planning unifié dans la 
Suite ........ : table YPLANNING
Mots clefs ... : PGPLANNING
*****************************************************************}
function DeletePGYPL (Const GuidOri, GuidYRS : string): integer;
var
GUID, St : string;
QRechGuid : TQuery;
i : integer;
begin
result:= 0;
if not GestionYPlanning then
   exit;
   
try
   if (GuidOri<>'') then
      result:= DeleteYPL (cStTablePaie, GuidOri)
   else
      begin
      St:= 'SELECT YPL_GUIDORI FROM YPLANNING WHERE YPL_GUIDYRS="'+GuidYRS+'"';
      QRechGuid:= OpenSql(St, TRUE);
      if (not QRechGUID.Eof) then
         begin
         QRechGUID.First;
         for i:=0 to QRechGuid.RecordCount-1 do
             begin
             GUID:= QRechGUID.Fields[0].AsString;
             result:= DeleteYPL (cStTablePaie, GUID);
             QRechGUID.Next;
             end;
         end;
      Ferme (QRechGuid);
      end;
finally
   if (GuidOri='') then
      Ferme (QRechGuid);
   end;
end;


end.
 