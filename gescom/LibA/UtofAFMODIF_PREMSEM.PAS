{***********UNITE*************************************************
Auteur  ...... : PL
Créé le ...... : 18/11/2002
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : AFMODIF_PREMSEM ()
Mots clefs ... : TOF;AFMODIF_PREMSEM
*****************************************************************}
Unit UtofAFMODIF_PREMSEM ;

Interface

Uses
{$IFDEF EAGLCLIENT}
      eMul,MaineAGL,
{$ELSE}
       HDB, mul, db,dbTables,Fe_Main,
{$ENDIF}

      sysUtils,Controls, Classes, ParamSOc, HStatus, HCtrls, HEnt1, Ent1, UTOB, UTOF, HmsgBox, AffaireUtil;

Type
  TOF_AFMODIF_PREMSEM = Class (TOF)
    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnClose                  ; override ;
  Private
    PremSem : string;
    Erreur : Boolean;
    procedure AppelChangePremSem;
    procedure ChangePremSemActivite;
    procedure ChangePremSemAfCumul;
  end ;

    
Procedure AFLanceFiche_Modif_PremSem;


Implementation

procedure TOF_AFMODIF_PREMSEM.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_AFMODIF_PREMSEM.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_AFMODIF_PREMSEM.OnUpdate ;
var
text, titre : string;
RepVal : TIOErr;
begin
  Inherited ;
If (GetControlText ('PREMSEMNEW') = '') or
   (GetControlText ('PREMSEMNEW') = PremSem) then
  begin
    PGIInfo ('Vous n''avez pas changé de valeur ou valeur vide', titre);
    Lasterror := 1;
    Erreur := True;
    exit;
  end
else
  begin
    text := 'Confirmez-vous le changement de l''option de première semaine';
    If (PGIAsk (text,titre) <> mrYes) then exit;

    SetControlText ('SUIVI', 'Mise à jour table en cours');

      // dans tous les cas on détruit les enrgt existant dans la table
    RepVal := Transactions (AppelChangePremSem, 2);
    Case RepVal of
        oeOk :
            BEGIN
            PGIInfo ('La modification de l''option de première semaine s''est déroulée avec succés.', titre);
            END;
        oeUnknown, oeSaisie :
            BEGIN
            MessageAlerte('ATTENTION : Changement d''option de première semaine non effectué !');
            Exit;
            END ;
    END ;
  end;
end ;

procedure TOF_AFMODIF_PREMSEM.AppelChangePremSem;
var
Semaine53Old : integer;
begin
  Semaine53Old := GetParamsoc ('SO_PREMIERESEMAINE');
  try
    // stocke nouvelle valeur du paramsoc
    SetParamsoc ('SO_PREMIERESEMAINE', GetControlText('PREMSEMNEW'));


    // processus de changement de gestion de première semaine dans l'activite    
    ChangePremSemActivite;


    // processus de changement de gestion de première semaine dans ...
    ChangePremSemAfCumul;


    // C'EST ICI QU'IL FAUT RAJOUTER LES MODIFS DANS LES TABLES QUI CONTIENNENT LA ZONE NUM de SEMAINE


    // on recharge les paramsoc dans la variable globale si tout s'est bien passé
    if (V_PGI.IoError = oeOk) then
      begin
      ChargeSocieteHalley;
      V_PGI.Semaine53 := GetParamsoc('SO_PREMIERESEMAINE');
      end
    else
      raise ERangeError.Create('');

  except
    // en cas d'erreur, on revient en arrière dans la variable globale
    V_PGI.Semaine53 := Semaine53Old;
  end;

end;

procedure TOF_AFMODIF_PREMSEM.ChangePremSemActivite;
var
TobRes,TobACT,TobDEt : TOB;
QQ1,QQ2:Tquery;
ii,jj   : longint;
NSemaine : integer;
begin
TobRes := nil; TobAct := nil; QQ1 := nil;
try
SourisSablier;
InitMove (3,'');

if Not(ExisteSQL('SELECT ACT_TYPEACTIVITE FROM ACTIVITE')) then exit ; // rien dans la table activite dans la base, on sort...
MoveCur(False) ;


if V_PGI.IoError<>oeOk then exit;

TobRes := Tob.Create ('les ress', Nil, -1);
  // on traite l'activité par ressource pour éviter de récupérer trop d'enrgt à la fois
QQ1 := OpenSql ('SELECT ARS_RESSOURCE FROM RESSOURCE', true);

if not (QQ1.eof) then
  begin
    TOBRes.LoadDetailDB ('RESSOURCE', '', '', QQ1, False);
    TOBDet := TOB.Create ('Les Ressources', TobRes, -1);
    TOBDet.LoadFromSt ('ARS_RESSOURCE;', ';', ';', ';');

    InitMove (TobRes.Detail.count, '');
    for ii := 0 to TobRes.Detail.count-1 do
      BEGIN
      if V_PGI.IoError<>oeOk then exit;
      QQ2 := nil;
      try
      // PL le 12/05/03 : changement de clé de la table ACTIVITE
//      QQ2 := Opensql('SELECT ACT_TYPEACTIVITE,ACT_AFFAIRE,ACT_RESSOURCE,ACT_DATEACTIVITE,ACT_FOLIO,ACT_TYPEARTICLE,ACT_NUMLIGNE,ACT_SEMAINE FROM ACTIVITE WHERE ACT_RESSOURCE="'
      QQ2 := Opensql('SELECT ACT_TYPEACTIVITE,ACT_AFFAIRE,ACT_NUMLIGNEUNIQUE,ACT_RESSOURCE,ACT_DATEACTIVITE,ACT_FOLIO,ACT_TYPEARTICLE,ACT_NUMLIGNE,ACT_SEMAINE FROM ACTIVITE WHERE ACT_RESSOURCE="'
            + TobRes.detail[ii].GetValue('ARS_RESSOURCE')+'"',True);
      if Not (QQ2.Eof) then
        begin
          try
          TobAct := Tob.Create ('les Act', Nil, -1);
          TobAct.LOadDetailDb ('ACTIVITE', '', '', QQ2, False);
          For jj :=0 to TobAct.Detail.count-1 do
            begin
            TobDet := TObAct.detail[jj];
            NSemaine := NumSemaine(TOBDet.Getvalue('ACT_DATEACTIVITE'));
            TobDet.Putvalue ('ACT_SEMAINE', NSemaine);
            end;

//          if (Not TobAct.InsertOrUpdateDb (false)) then // PL le 12/05/03 : on ne doit pas faire d'insert si on sait que c'est forcement un update !!!
          if (Not TobAct.UpdateDb) then
            V_PGI.IoError := oeUnknown;

          finally
          TobAct.free;
          TobAct := nil;
          end;
        end;
      finally
        Ferme (QQ2);
      end;

      MoveCur(False) ;
      END;  // boucle for
  end;

finally
  Ferme (QQ1);
  TobRes.free;
  FiniMove;
  SourisNormale;
end;

end;

procedure TOF_AFMODIF_PREMSEM.ChangePremSemAfCumul;
var
TobTypeDates, TobAFC, TobDet : TOB;
QQ1, QQ2 : Tquery;
ii, jj : longint;
NSemaine : integer;
sReq2 : string;
begin
TobTypeDates := nil; TobAFC := nil; QQ1 := nil;
try
SourisSablier;
InitMove (3,'');

if Not (ExisteSQL ('SELECT ACU_TYPEAC FROM AFCUMUL')) then exit ; // rien dans la table afcumul dans la base, on sort...
MoveCur (False) ;


if (V_PGI.IoError <> oeOk) then exit;

TobTypeDates := Tob.Create ('les TypeDate', Nil, -1);
  // on traite l'afcumul par Type/date pour éviter de récupérer trop d'enrgt à la fois
QQ1 := OpenSql ('SELECT ACU_TYPEAC, ACU_DATE FROM AFCUMUL GROUP BY ACU_TYPEAC, ACU_DATE', true);

if not (QQ1.eof) then
  begin
    TobTypeDates.LoadDetailDB ('AFCUMUL', '', '', QQ1, False);

    InitMove (TobTypeDates.Detail.count, '');
    for ii := 0 to TobTypeDates.Detail.count-1 do
      BEGIN
      if V_PGI.IoError<>oeOk then exit;
      QQ2 := nil;
      try
      sReq2 := 'SELECT ACU_TYPEAC,ACU_DATE,ACU_TIERS,ACU_AFFAIRE,ACU_NUMECLAT,ACU_RESSOURCE,'
                    + 'ACU_TYPEARTICLE,ACU_CODEARTICLE,ACU_SEMAINE FROM AFCUMUL WHERE '
                    + 'ACU_TYPEAC="' + TobTypeDates.detail[ii].GetValue('ACU_TYPEAC') + '" '
                    + 'AND ACU_DATE="' + UsDateTime(TobTypeDates.detail[ii].GetValue('ACU_DATE')) + '"';
      QQ2 := Opensql (sReq2, True);
      if Not (QQ2.Eof) then
        begin
          try
          TobAFC := Tob.Create ('les Afc', Nil, -1);
          TobAFC.LoadDetailDb ('AFCUMUL', '', '', QQ2, False);
          For jj := 0 to TobAFC.Detail.count - 1 do
            begin
            TobDet := TobAFC.detail[jj];
            NSemaine := NumSemaine (TobDet.Getvalue('ACU_DATE'));
            TobDet.Putvalue ('ACU_SEMAINE', NSemaine);
            end;

          if (Not TobAFC.InsertOrUpdateDb (false)) then
            V_PGI.IoError := oeUnknown;

          finally
          TobAFC.free;
          TobAFC := nil;
          end;
        end;
      finally
        Ferme (QQ2);
      end;

      MoveCur (False) ;
      END;  // boucle for
  end;

finally
  Ferme (QQ1);
  TobTypeDates.free;
  FiniMove;
  SourisNormale;
end;

end;

procedure TOF_AFMODIF_PREMSEM.OnLoad ;
begin
  Inherited ;
end ;

procedure TOF_AFMODIF_PREMSEM.OnArgument (S : String ) ;
begin
  Inherited ;
Erreur:=False;
PremSem := GetParamSoc('SO_PREMIERESEMAINE');
SetControltext ('PREMSEMOLD', PremSem);
end ;

procedure TOF_AFMODIF_PREMSEM.OnClose ;
begin
  Inherited ;
If Erreur then begin
   LastError:=(-1);
   LastErrorMsg:='';
   Erreur:=false;
   exit;     // il y eu une erreur, on reste sur la même fiche
   end;
end ;


Procedure AFLanceFiche_Modif_PremSem;
begin
if ToutSeulAff  then  exit; // quelqu'un d'autre travaille sur la base
try
AGLLanceFiche ('AFF','AFMODIF_PREMSEM','','','');
finally
Bloqueur ('AffToutSeul',False);
end;
end;

Initialization

  registerclasses ( [ TOF_AFMODIF_PREMSEM ] ) ;
end.
