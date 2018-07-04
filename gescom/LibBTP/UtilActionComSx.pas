unit UtilActionComSx;

interface
uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  Ent1,Hent1,ParamSoc,
  IniFiles;


type

	TmodeActionComSx = (TmaSxImpInit,TmaSxExpInit,TmaSxImport,TmaSxExport);
         
	TEchangeComSx = class
  	private
      NomFichierCmd : string;
      FichierATraiter : string;
      DateMax,DateDebut,DateFin : string;
      Envoye : string;
      ActionComSx : TModeActionComSx;
			procedure CreationFichierCmds;
    	procedure GenereIniRecup;
    public
    	procedure Init;
    	procedure ExecuteAction(LActionComSx : TModeActionComSx ; Fichier: string; DatePEC : string='';DateD : string = ''; DateF : string=''; DejaEnvoye : string='-');
  end;

implementation

procedure TEchangeComSx.CreationFichierCmds;
var RepertoireExe : String;
		ChaineCmd : string;
begin

	// Definition du nom de fichier de commande .INI
	if ActionComSx = TmaSxImpInit then NomFichierCmd := ExtractFiledir (FichierATraiter)+'\RECUPINIT'
  else if ActionComSx = TmaSxImport then NomFichierCmd := ExtractFiledir (FichierATraiter)+'\IMPORT'
  else if ActionComSx = TmaSxExport then NomFichierCmd := ExtractFiledir (FichierATraiter)+'\EXPORT';
  //
	GenereIniRecup;

end;

procedure TEchangeComSx.GenereIniRecup;
var RepertoireCmd: string;
		FichierUniquement : string;
    FichierCmd : TIniFile;
    TranfertVers : string;
    Email : string;
begin
	Email := GetparamSocSecur('SO_CPRDEMAILCLIENT','');
  TranfertVers := GetParamSocSecur ('SO_CPLIENGAMME','S1');
  FichierUniquement := ExtractFileName (FichierATraiter);
  RepertoireCmd := ExtractFiledir (NomFichierCmd);
  FichierCmd := TIniFile.create(NomFichierCmd+'.INI');
  if ActionComSx = TmaSxImpInit then // Recup du dossier complet
  begin
    FichierCmd.writeString('COMMANDE','REPERTOIRE',RepertoireCmd);
    FichierCmd.writeString('COMMANDE','NOMFICHIER',FichierUniquement);
    FichierCmd.writeString('COMMANDE','MAIL',Email);
  end else if ActionComSx = TmaSxImport then // recup du dossier
  begin
    FichierCmd.writeString('COMMANDE','REPERTOIRE',RepertoireCmd);
    FichierCmd.writeString('COMMANDE','NOMFICHIER',FichierUniquement);
    FichierCmd.writeString('COMMANDE','MAIL',Email);
  end else if ActionComSx = TmaSxExport then // envoie des journaux
  begin
    FichierCmd.writeString('COMMANDE','NOMFICHIER',FichierATraiter);
    FichierCmd.writeString('COMMANDE','NATURETRANSFERT','JRL');
    FichierCmd.writeString('COMMANDE','TRANSFERTVERS',TranfertVers);
    FichierCmd.writeString('COMMANDE','MAIL',Email);
    FichierCmd.writeString('COMMANDE','FORMAT','ETE');
    FichierCmd.writeString('COMMANDE','EXERCICE','');
    FichierCmd.writeString('COMMANDE','DATEECR1','01/01/1900');
    FichierCmd.writeString('COMMANDE','DATEECR2','01/01/2099');
    if DateMax = '' then Datemax := DateToStr (Now);
    FichierCmd.writeString('COMMANDE','DATEARRET',DateMax);
    FichierCmd.writeString('COMMANDE','EXCLURE','');
    if DateDebut <> '' then
    begin
      FichierCmd.writeString('COMMANDE','DATEECR1',DateDebut);
      if DateFin='' then DateFIn := DateToStr(V_PGI.DateEntree);
      FichierCmd.writeString('COMMANDE','DATEECR2',DateFin);
    end;
    if Envoye <> '' then
    begin
      FichierCmd.writeString('COMMANDE','EXPORTE',Envoye);
    end;
  end;
  FichierCmd.free;
end;



procedure TEchangeComSx.ExecuteAction (LActionComSx : TModeActionComSx ; Fichier: string; DatePEC : string='';DateD : string = ''; DateF : string=''; DejaEnvoye : string='-');
var RepertoireExe,ChaineCmd : string;
		ChaineImpExp : string;
begin
	ActionComSx := LActionComSx;
	FichierATraiter := Fichier;
  Datemax := DatePEc;
  DateDebut := DateD;
  DateFin := DateF;
  if DejaEnvoye='X' then Envoye:='TRUE'
  else if DejaEnvoye='-' then Envoye:='FALSE'
  else Envoye := '';
  CreationFichierCmds;

	if ActionComSx = TmaSxExport then
  begin
  	ChaineImpExp := 'EXPORT'
  end else
  begin
  	ChaineImpExp := 'IMPORT';
  end;
  RepertoireExe := ExtractFileDir (Application.ExeName)+'\COMSX.exe';
  ChaineCmd := RepertoireExe+' /USER='+V_PGI.UserLogin+ ' /PASSWORD='+V_PGI.PassWord+
  						 ' /DOSSIER='+V_PGI.DBName+ ' /INI='+NomFichierCmd+'.INI;'+ChaineImpExp+';Minimized';

  FileExecAndWait (ChaineCmd);
  //
  DeleteFile (NomFichierCmd+'.INI');
end;

procedure TEchangeComSx.Init;
begin
	NomFichierCmd := '';
  FichierATraiter := '';
end;

end.
