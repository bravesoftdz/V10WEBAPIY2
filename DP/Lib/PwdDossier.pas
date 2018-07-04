unit PwdDossier;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  StdCtrls, Buttons, ExtCtrls, Hctrls, hmsgbox, HTB97, PGIExec, PGIAppli,
  GalOutil;

type
  TFPwdDossier = class(TForm)
    FPassWord: TEdit;
    Panel1: TPanel;
    BValider: TToolbarButton97;
    BFerme: TToolbarButton97;
    BAide: TToolbarButton97;
    TPassword: THLabel;
    procedure BValiderClick(Sender: TObject);
  private
    { Déclarations privées }
    GoodPassword : String;
  public
    { Déclarations publiques }
  end;

function VerifPwdDossier(nodoss: String; permissif: Boolean=True): Boolean ;
function VerifPwdSociete: Boolean;
function GetPwdDossier(nodoss: String): String;
function CheckPwdDossier(nodoss, pwd: String): Boolean;
function RecupPwdDossierFromLgcde: String;
function ToutesApplisProtected(nodoss: String; LstAppProtec: HTStringList): Boolean;

///////////// IMPLEMENTATION /////////////
implementation

{$R *.DFM}

uses
{$IFDEF EAGLCLIENT}
   UTob,
{$ELSE}
   {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF}
{$ENDIF}
   LicUtil;


function VerifPwdDossier(nodoss: String; permissif: Boolean=True) : Boolean ;
// si permissif, renvoit True si seules quelques applis sont protégées
var F: TFPwdDossier;
    pwd : String;
begin
  Result := True;
  // Renvoit ok si mot de passe non global
  if permissif and (Not ToutesApplisProtected(nodoss, Nil)) then exit;

  // Si pas de mot de passe
  pwd := GetPwdDossier(nodoss);
  if pwd='' then exit;

  // Si mot de passe, le faire saisir
  F:=TFPwdDossier.Create(Application) ;
  F.Caption := 'Mot de passe du dossier ' + nodoss + ' :';
  F.GoodPassword := DecryptageSt(pwd);
  try
    Result:=(F.ShowModal=mrOK) ;
  finally
    F.Free ;
    end ;
end;


function VerifPwdSociete: Boolean;
// Vérif du mot de passe en connexion directe à une soc
var codsoc, nodoss : String;
    Q : TQuery;
begin
  Result := True;
  codsoc := '';
  // en connexion directe, V_PGI.NoDossier contient 000000, donc non utilisable !
  // => on va récupérer un n° de dossier à partir du code soc ...
  Q := OpenSQL('SELECT SO_SOCIETE FROM SOCIETE', True,-1,'',true);
  if Not Q.Eof then codsoc := Q.FindField('SO_SOCIETE').AsString;
  Ferme(Q);
  if codsoc='' then exit;
  nodoss := '';
  Q := OpenSQL('SELECT DOS_NODOSSIER FROM DOSSIER WHERE DOS_SOCIETE="'+codsoc+'"', True,-1,'',true);
  if Not Q.Eof then nodoss := Q.FindField('DOS_NODOSSIER').AsString;
  Ferme(Q);
  if nodoss='' then exit;
  // vérif du pwd par rapport à ce n° de dossier
  Result := VerifPwdDossier(nodoss);
end;


function GetPwdDossier(nodoss: String): String;
// Retourne le password crypté
var Q: TQuery;
begin
  Result := '';
  // petite vérif, même si cantonnée à la base en cours... (quid si DOSSIER redirigée ?)
  // ou alors tester ChampPhysiqueExiste(nombase.dbo.nomtable...) mais ça fait un SELECT
  if ChampToNum('DOS_PASSWORD')<0 then exit;

  Q := OpenSQL('SELECT DOS_PASSWORD FROM ##DP##.DOSSIER WHERE DOS_NODOSSIER="'+nodoss+'"', True,-1,'',true);
  if Not Q.Eof then Result := Q.FindField('DOS_PASSWORD').AsString;
  Ferme(Q);
end;


function CheckPwdDossier(nodoss, pwd: String): Boolean;
// comparaison des password, en crypté
begin
  Result := ( GetPwdDossier(nodoss) = pwd );
end;


function RecupPwdDossierFromLgcde: String;
// Récupère le pwddossier passé en ligne de commande ou via PGIExec
var St, Nom, Valeur : String;
    i : Integer;
begin
  Result := '';
  // privilégie la ligne de commande
  for i:=1 to ParamCount do
    begin
    St := ParamStr(i);
    Nom := UpperCase(Trim(ReadTokenPipe(St, '=')));
    Valeur := UpperCase(Trim(St));
    if Nom = '/PWDDOSSIER' then
      begin
      Result := Valeur;
      exit;
      end;
    end;
  // sinon PGIExec
  if PGIApp<>Nil then
    Result := PGIApp.GetParam('General', 'PwdDossier', '');
end;


procedure TFPwdDossier.BValiderClick(Sender: TObject);
begin
  if FPassWord.Text <> GoodPassword then
    PGIInfo('Mot de passe incorrect !')
  else
    modalResult := mrOK;
end;


function ToutesApplisProtected(nodoss: String; LstAppProtec: HTStringList): Boolean;
// - retourne True si toutes les applis sont protégées par le mot de passe éventuel
//            False sinon
// - renseigne un Tstringlist de la liste des exe concernés par la protection
//   (les exe sont stockés en majuscules, avec l'extension .EXE)
var Q: TQuery;
    tmp, uneappli : String;
    i : Integer;
begin
  Result := True;
  // petite vérif, même si cantonnée à la base en cours... (quid si DOSSIER redirigée ?)
  // ou alors tester ChampPhysiqueExiste(nombase.dbo.nomtable...) mais ça fait un SELECT
  if ChampToNum('DOS_PASSWORD')<0 then exit;

  if (V_Applis=Nil) or (V_Applis.Applis.Count=0) then exit;

  Q := OpenSQL('SELECT DOS_PWDGLOBAL FROM ##DP##.DOSSIER WHERE DOS_NODOSSIER="'+nodoss+'"', True,-1,'',true);
  if Not Q.Eof then
    begin
    Result := (Q.FindField('DOS_PWDGLOBAL').AsString='X');
    tmp:=DonnerListeAppliProtec (NoDoss);
    if LstAppProtec<>Nil then
      begin
      While tmp<>'' do
        begin
        uneappli := ReadTokenSt(tmp);
        // pour chaque éxé...
        for i := 0 to V_Applis.Applis.Count - 1 do
          begin
          if UpperCase(TPGIAppli(V_Applis.Applis[i]).Nom) = UpperCase(uneappli)+'.EXE' then
            begin
            LstAppProtec.Add(UpperCase(uneappli+'.EXE'));
            break;
            end;
          end;
        end;
      end;
    end;
  Ferme(Q);
end;


end.
