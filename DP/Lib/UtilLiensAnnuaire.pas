unit UtilLiensAnnuaire;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
   HCtrls;

procedure SupprimeLienAnnuaire(Guidperdos, Guidper, typedos,
  fct: String; intseulmt: Boolean=FALSE);

{+GHA - 11/2007 FQ 11860}
procedure ModifieLienAnnuaire(Guidperdos, Guidper, typedos,
  fct: String);
{-GHA - 11/2007 FQ 11860}

//////////// IMPLEMENTATION /////////////
implementation

procedure SupprimeLienAnnuaire(Guidperdos, Guidper, typedos,
  fct: String; intseulmt: Boolean=FALSE);
var
  sql : string;
begin
  if (Guidperdos='') or (Guidper='') then exit;

  // suppression du lien en cours (la clé ANNULIEN est :
  // ANL_CODEPERDOS,ANL_TYPEDOS,ANL_NOORDRE,ANL_FONCTION,ANL_CODEPER
  // mais NOORDRE est toujours à 1)
  if not intseulmt then
    ExecuteSQL('delete from ANNULIEN where ANL_GUIDPERDOS="'+Guidperdos
     +'" and ANL_TYPEDOS="'+typedos+'" and ANL_FONCTION="'+fct+'"'
     +' and ANL_GUIDPER="'+Guidper+'"');

  // suppression du lien INTERVENANT (loi NRE)
  // s'il est tout seul et non lié à un dossier juridique
  if Not ExisteSQL('select 1 from ANNULIEN where ANL_GUIDPERDOS="'+guidperdos
   +'" AND ANL_GUIDPER="'+Guidper+'" AND ANL_FONCTION<>"INT"') // pas de lien autre que intervenant
  // et vérifie que l'éventuel lien INT n'est pas utilisé par un dossier juridique
  and Not ExisteSQL('select 1 from ANNULIEN where ANL_GUIDPERDOS="'+Guidperdos
   +'" AND ANL_GUIDPER="'+GuidPer+'" AND ANL_FONCTION="INT" AND ANL_CODEDOS<>"&#@"') then
    ExecuteSQL('delete from ANNULIEN where ANL_GUIDPERDOS="'+Guidperdos
     +'" AND ANL_GUIDPER="'+Guidper+'" AND ANL_FONCTION="INT"');

  {+GHA - 11/2007 FQ 11860}
  //Si suppression du lien pour un dossier de type Filiale.
  //on supprime le lien du dossier de type Tête de groupe.
  if fct = 'FIL' then
  begin
    sql := 'DELETE FROM ANNULIEN'+
         ' WHERE ANL_GUIDPERDOS = "'+Guidper+'"'+
         ' AND ANL_GUIDPER = "'+Guidperdos+'"'+
         ' AND ANL_TYPEDOS = "'+typedos+'"'+
         ' AND ANL_NOORDRE = 1'+
         ' AND ANL_FONCTION = "TIF"';
    ExecuteSQL(sql);
  end;
  {-GHA - 11/2007 FQ 11860}

end;
///////////////////////////////////////////////////////////////////////////////
{+GHA - 11/2007 FQ 11860}
procedure ModifieLienAnnuaire(Guidperdos, Guidper, typedos, fct: String);
var
  sql : string;
begin
  //Mise à jour la table ANNULIEN dans le cas
  //ou l'on supprime le dossier de type Tête de groupe
  //à partir d'un dossier de type fifiale.
  //Mise à jour du champ ANL_GRPFISAL. = FALSE
  if (Guidperdos='') or (Guidper='') then
    exit;

  sql := 'SELECT 1 FROM ANNULIEN'+
         ' WHERE ANL_GUIDPERDOS = "'+Guidperdos+'"'+
         ' AND ANL_GUIDPER = "'+Guidper+'"'+
         ' AND ANL_TYPEDOS = "'+typedos+'"'+
         ' AND ANL_NOORDRE = 1'+
         ' AND ANL_FONCTION = "'+fct+'"';

  if not ExisteSQL(sql) then
    exit;

  //Mise à jour du lien
  sql := 'UPDATE ANNULIEN SET ANL_GRPFISCAL = "-"'+
         ' WHERE ANL_GUIDPERDOS = "'+Guidperdos+'"'+
         ' AND ANL_GUIDPER = "'+Guidper+'"'+
         ' AND ANL_TYPEDOS = "'+typedos+'"'+
         ' AND ANL_NOORDRE = 1'+
         ' AND ANL_FONCTION = "'+fct+'"';

  ExecuteSQL(Sql);
end;
{-GHA - 11/2007 FQ 11860}


end.
