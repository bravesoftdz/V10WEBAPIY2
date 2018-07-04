{***********UNITE*************************************************
Auteur  ...... :
Créé le ...... : 13/11/2006
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : UTOF_PGMULPOPMODULES ()
Mots clefs ... : TOF;UTOF_PGMULPOPMODULES
*****************************************************************}
unit UTOF_PG_MULPOPMODULES;

interface

uses StdCtrls,
  Controls,
  Classes,
{$IFNDEF EAGLCLIENT}
  db,
{$IFNDEF DBXPRESS}dbtables, {$ELSE}uDbxDataSet, {$ENDIF}
  mul,
{$ELSE}
  eMul,
  uTob,
{$ENDIF}
  forms,
  sysutils,
  ComCtrls,
  HCtrls,
  HEnt1,
  HMsgBox,
  UTOF;

type
  TOF_PGMULPOPMODULES = class(TOF)
    procedure OnNew; override;
    procedure OnDelete; override;
    procedure OnUpdate; override;
    procedure OnLoad; override;
    procedure OnArgument(S: string); override;
    procedure OnDisplay; override;
    procedure OnClose; override;
    procedure OnCancel; override;
  private
    procedure AppliquerPop(Sender: TObject);
  end;

implementation
uses P5Def, HTB97, PgOutils2, pgpopuloutils;


procedure TOF_PGMULPOPMODULES.OnNew;
begin
  inherited;
end;

procedure TOF_PGMULPOPMODULES.OnDelete;
begin
  inherited;
end;

procedure TOF_PGMULPOPMODULES.OnUpdate;
begin
  inherited;
end;

procedure TOF_PGMULPOPMODULES.OnLoad;
var
  nodossier, pred: string; // , sql1
  populok: boolean;
begin
  inherited;
  populok := canusepopulation('PAI');
  if populok then
  begin
    // Pt1 recherche du prédéfini des populations PAI à afficher
    pred := GetPredefiniPopulation('PAI');
    nodossier := PgrendNodossier();
    // recherche des codes pop prédéfini DOSSIER
{    pred := 'DOS';
    sql1 := 'select PPC_POPULATION,PPO_CODEPOP from ORDREPOPULATION ,CODEPOPULATION where PPC_CODEPOP = PPO_CODEPOP'+
    ' and PPO_PREDEFINI= "DOS" and PPO_NODOSSIER = "'+nodossier+'" and PPO_TYPEPOP LIKE "%PAI%"';
    if not existesql(sql1) then
    begin
       // recherche des codes pop prédéfini STANDARD
       sql1 := 'select PPC_POPULATION,PPO_CODEPOP from ORDREPOPULATION,CODEPOPULATION where PPC_CODEPOP = PPO_CODEPOP'+
       ' and PPO_PREDEFINI= "STD" and PPO_TYPEPOP LIKE "%PAI%" ';
       pred := 'STD';
       if not existesql(sql1) then
       begin
         // recherche des codes pop prédéfini CEGID
         sql1 := 'select PPR_POPULATION, PPO_CODEPOP from ORDREPOPULATION, CODEPOPULATION where PPC_CODEPOP = PPO_CODEPOP'+
         ' and PPO_PREDEFINI= "CEG" and PPO_TYPEPOP LIKE "%PAI%" ';
         pred := 'CEG';
       end;
    end; }
    SetControlText('XX_WHERE', '##PPC_PREDEFINI## PPC_PREDEFINI = "' + pred + '" and PPO_TYPEPOP LIKE "%PAI%"');
  end else begin
    setcontrolText('XX_WHERE', 'PPO_TYPEPOP LIKE "%XXX%"');
    PgiInfo('Le type de population "PAI" n''est pas valide.', Ecran.caption);
  end;
end;

procedure TOF_PGMULPOPMODULES.OnArgument(S: string);
var Q: TQuery;
  BtnValidMul: TToolbarButton97;
begin
  inherited;
  if Assigned(MonHabilitation) then
    if (MonHabilitation <> nil) and (MonHabilitation.Active) then

    begin
      Q := OpenSQL('SELECT PPO_LIBELLE FROM CODEPOPULATION WHERE PPO_CODEPOP="' + MonHabilitation.CodePop + '"', TRUE);
      if not Q.EOF then
      begin
        SetControlVisible('LBLHABILITATION', True);
        SetControlText('LBLHABILITATION', GetControlText('LBLHABILITATION') + ' ' + Q.FindField('PPO_LIBELLE').AsString);
      end;
      if MonHabilitation.LibellePop <> '' then
        SetControlText('LBLPOP', GetControlText('LBLPOP') + MonHabilitation.LibellePop)
      else SetControlText('LBLPOP', 'Aucune population sélectionnée.');
      Ferme(Q);
    end;
  BtnValidMul := TToolbarButton97(GetControl('BOuvrir'));
  if BtnValidMul <> nil then BtnValidMul.OnClick := AppliquerPop;
end;

procedure TOF_PGMULPOPMODULES.OnClose;
begin
  inherited;
end;

procedure TOF_PGMULPOPMODULES.OnDisplay();
begin
  inherited;
end;

procedure TOF_PGMULPOPMODULES.OnCancel();
begin
  inherited;
end;

procedure TOF_PGMULPOPMODULES.AppliquerPop(Sender: TObject);
var i: Integer;
  CodePop, Population, St, St1, St2, LibellePop, LePlus: string; // , MonSql
begin
  with TFMul(Ecran) do
  begin
    if (FListe.NbSelected = 0) and (not FListe.AllSelected) then
    begin
      PgiInfo('Aucune population sélectionnée.', Ecran.caption);
      MonHabilitation.SqlPop := '';
      MonHabilitation.LibellePop := '';
      PgVideHabilitation();
      PgRendHabilitation ();
    end
    else
    begin
      for i := 0 to FListe.NbSelected - 1 do
      begin
        FListe.GotoLeBOOKMARK(i);
{$IFDEF EAGLCLIENT}
        Q.TQ.Seek(FListe.Row - 1);
{$ENDIF}
        CodePop := TFmul(Ecran).Q.FindField('PPC_CODEPOP').AsString;
        Population := TFmul(Ecran).Q.FindField('PPC_POPULATION').AsString;
        if (CodePop <> '') and (Population <> '') then
        begin
          if MonHabilitation.Habilitation then
          begin
            PGrechwhere(Codepop, population, False, st1, st2);
            st := st + ' AND ' + st1;
            if LibellePop <> '' then LePlus := ' et '
            else LePlus := '';
            LibellePop := LibellePop + LePlus + TFmul(Ecran).Q.FindField('PPC_LIBELLE').AsString;
          end
          else
          begin
             PgRendHabilitation(codePop, Population, TFmul(Ecran).Q.FindField('PPC_LIBELLE').AsString);
             LibellePop := TFmul(Ecran).Q.FindField('PPC_LIBELLE').AsString;
          end;
        end;
      end;
      // Fin de la boucle sur la recherche du sql des codes population
      if st <> '' then // on applique des critères en plus de ceux de l'habilitation
      begin
        MonHabilitation.SqlPop := st;
        MonHabilitation.LibellePop := LibellePop;
        if (not MonHabilitation.Active) then MonHabilitation.Active := true;
      end;
    end;
  end;
  TFMul(Ecran).BAnnulerClick(nil);
end;

initialization
  registerclasses([TOF_PGMULPOPMODULES]);
end.

