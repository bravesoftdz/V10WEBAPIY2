unit UtilModeleConso;

interface

Uses StdCtrls,
     Controls,
     Classes,
{$IFNDEF EAGLCLIENT}
     db,
     {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF}
     mul,
{$else}
     eMul,
{$ENDIF}
     forms,
     LookUp,
     sysutils,
     ComCtrls,
     HCtrls,
     HEnt1,
     HMsgBox,
     UTOB ;

function LoadModele (TOBModele : TOB; Modele,TypeRessource : string) : boolean;
function IsModeleExist (Modele,TypeRessource : string) : boolean;
function GetLibelleModele (Modele,TypeRessource : string) : string;


implementation

function LoadModele (TOBModele : TOB; Modele,TypeRessource : string) : boolean;
var QQ : TQuery;
    req : string;
begin
  TOBModele.ClearDetail;
  TOBModele.InitValeurs;
  //
  QQ := OpenSql ('SELECT * FROM BTMODELSHEBDO WHERE BMS_TYPERESSOURCE="'+TypeRessource+'" AND BMS_CODEMODELE="'+Modele+'"',True);
  result := not QQ.eof;
  if result then
  begin
    TOBModele.selectDb ('',QQ);
    ferme (QQ);
    req :='SELECT *,BNP_TYPERESSOURCE,BNP_LIBELLE,GA_TYPEARTICLE,GA_DPA,GA_DPR,GA_PVHT,GA_QUALIFUNITEVTE,GA_FOURNPRINC FROM BTDETMODELSHEB '+
          'LEFT JOIN ARTICLE ON BSM_ARTICLE=GA_ARTICLE '+
          'LEFT JOIN NATUREPREST ON GA_NATUREPRES=BNP_NATUREPRES '+
          'WHERE BSM_TYPERESSOURCE="'+TypeRessource+'" AND BSM_CODEMODELE="'+Modele+'" ORDER BY BSM_NUMORDRE';
    QQ := OpenSql (Req,True);
    TOBModele.LoadDetailDB ('BTDETMODELSHEB','','',QQ,false);
  end;
  ferme (QQ);
end;

function GetLibelleModele (Modele,TypeRessource : string) : string;
var QQ : TQuery;
begin
  result := '';
  QQ := OpenSql ('SELECT BMS_LIBELLE FROM BTMODELSHEBDO WHERE BMS_TYPERESSOURCE="'+TypeRessource+'" AND BMS_CODEMODELE="'+Modele+'"',True);
  if not QQ.eof then result := QQ.FindField('BMS_LIBELLE').AsString;
  ferme (QQ);
end;

function IsModeleExist (Modele,TypeRessource : string) : boolean;
var QQ : TQuery;
begin
  QQ := OpenSql ('SELECT BMS_LIBELLE FROM BTMODELSHEBDO WHERE BMS_TYPERESSOURCE="'+TypeRessource+'" AND BMS_CODEMODELE="'+Modele+'"',True);
  result := not QQ.eof;
  ferme (QQ);
end;

end.
 