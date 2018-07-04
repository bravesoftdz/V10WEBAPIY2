{***********UNITE*************************************************
Auteur  ...... :
Créé le ...... : 01/06/2006
Modifié le ... : 10/05/2007
Description .. : Source TOF de la FICHE : AMEXPORT ()
Suite ........ : BTY FQ 19820 05/07 Mettre dans le .SAV les infos relatives à CRC200-10
Suite ........ : => nouveau paramètre d'appel
Mots clefs ... : TOF;AMEXPORT
*****************************************************************}
Unit AMEXPORT_TOF ;

Interface

Uses StdCtrls,
     Controls,
     Classes,
     Dialogs,
{$IFNDEF EAGLCLIENT}
     db,
     fe_main,
     {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF}
     mul,
{$else}
     eMul,
     MaineAGL,
     uTob,
{$ENDIF}
     forms,
     sysutils,
     ComCtrls,
     HCtrls,
     HEnt1,
     ImEnt,
     HMsgBox,
     HTB97,
     AmExport,
     FileCtrl,
     UTOF ;

Type
  TOF_AMEXPORT = Class (TOF)
    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnDisplay                ; override ;
    procedure OnClose                  ; override ;
    procedure OnCancel                 ; override ;
  private
    { Variables }
    Sauve : TSaveDialog;
    fCheminFichier : string;
    fEnrichi : Boolean;
    { Procédures}
    procedure OnCheminElipsisClick(Sender: TObject);
    function ControleCheminOK (Chemin : string) : boolean;
  end ;

procedure AMLanceFiche_AmExport (Action : string = ''); // FQ 18920

Implementation

const
 HM: array[0..1] of string =
 {0} ('Fichier non renseigné',
 {1}  'Nom de fichier incorrect' );

procedure AMLanceFiche_AmExport(Action : string = ''); // FQ 18920;
 begin
   AGLLanceFiche('AM', 'AMEXPORT','' , '', Action);  //'');
 end;

procedure TOF_AMEXPORT.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_AMEXPORT.OnDelete ;
begin
  Inherited ;
end ;

function TOF_AMEXPORT.ControleCheminOK(Chemin: string) : boolean ;
begin
   Result := True;
   fCheminFichier := Chemin;
   if (fCheminFichier = '') then
   begin
     Result := False;
     PGIBox(HM[0]);
   end else
   if (not DirectoryExists (ExtractFilePath(fCheminFichier)) )
   or (ExtractFileName(fCheminFichier)='') then
   begin
     Result := False;
     PGIBox(HM[1]);
   end else if ExtractFileExt(ExtractFileName(fCheminFichier))='' then
   begin
     fCheminFichier := fCheminFichier + '.sav';
     SetControlProperty('Chemin', 'Text', fCheminFichier);
   end;
end;

procedure TOF_AMEXPORT.OnUpdate ;
var   LExport : TAmImpExp;
begin
  Inherited ;

  if ControleCheminOK (GetControlText('Chemin')) then
     begin
     LExport := TAmImpExp.Create;
     LExport.OnInformation := nil;
     LExport.Exporte ( fCheminFichier ,True,True, fEnrichi);   // FQ 19820
     LExport.Free;
     PGIInfo('Export terminé.');
     // Ne pas revalider la fenêtre par inadvertance, cela relancerait l'export
     TToolbarButton97(GetControl('BValider')).Enabled := False;
     end;
end ;

procedure TOF_AMEXPORT.OnLoad ;
begin
  Inherited ;
end ;

procedure TOF_AMEXPORT.OnArgument (S : String ) ;
begin
  Inherited ;

  fEnrichi := (S = 'RI');  // FQ 19820
  if fEnrichi then
     SetControlCaption ('TABSHEET1', 'Exportation enrichie')
  else
     SetControlCaption ('TABSHEET1', 'Exportation standard');

  // Evènements
  THEdit(GetControl('Chemin')).OnElipsisClick  := OnCheminElipsisClick;

  // Préparation de l'export
  Sauve := TSaveDialog.Create(Ecran);
  Sauve.Filter := 'Fichier Texte (*.sav)|*.sav';
  Sauve.DefaultExt := 'SAV';
  Sauve.FilterIndex := 1;
  Sauve.Options := Sauve.Options + [ofOverwritePrompt, ofPathMustExist,
                   ofNoReadonlyReturn, ofNoLongNames] - [ofEnableSizing];
  Sauve.Title:= Ecran.Caption ;
  if ctxPCL in V_PGI.PGIContexte then Sauve.FileName:= 'IM2'+V_PGI.NoDossier+'.SAV'
  else Sauve.FileName := 'IMMO.SAV';

  {$IFNDEF EAGLCLIENT}
  if ctxPCL in V_PGI.PGIContexte then
     fCheminFichier := ChangeStdDatPath('$DOS\'+'IM2'+V_PGI.NoDossier+'.SAV',True)
  else
     fCheminFichier := 'C:\IMMO.SAV';
  SetControlProperty('Chemin', 'Text', fCheminFichier) ;
  {$ELSE}
  fCheminFichier := '';
  {$ENDIF}
end ;

procedure TOF_AMEXPORT.OnCheminElipsisClick(Sender: TObject);
begin
{$IFNDEF EAGLCLIENT}
ImDirDefault(Sauve, GetControlText('Chemin')) ;
{$ENDIF}
if Sauve.Execute then
   SetControlProperty('Chemin', 'Text', Sauve.FileName) ;
end;


procedure TOF_AMEXPORT.OnClose ;
begin
  Inherited ;
  if ASSIGNED(Sauve) then Sauve.Free;
end ;

procedure TOF_AMEXPORT.OnDisplay () ;
begin
  Inherited ;
end ;

procedure TOF_AMEXPORT.OnCancel () ;
begin
  Inherited ;
end ;

Initialization
  registerclasses ( [ TOF_AMEXPORT ] ) ;
end.
