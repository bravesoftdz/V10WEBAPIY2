unit TvaEdi;

interface

uses
    Shellapi,UTOF,HMsgBox,classes,Graphics,Controls,SysUtils,Dialogs,stdctrls,HStatus,comctrls,
    extctrls,WINDOWS,PGIEnv,DB,IniFiles,filectrl,DOC_Parser,{AglImportObject,}HPdfprev,HQuickrp,
    HCtrls, QRS1, HTB97,Vierge,HEnt1,buttons,UTOB,FE_Main, HPdfviewer,dbtables,
    Ent1,LIA_COMMUN, Paramsoc;

Function CreatFichierEdiTVA :Boolean;
Procedure EcritureRib (var F :TextFile);

//function TVACopierFichier (CheminSource, CheminDestination : string) : boolean;

implementation

Function CreatFichierEdiTVA :Boolean;
var
sSQL                :string;
QDos                :TQuery;
TOB_TEMPO           :TOB;
T                   :TOB;
Valrub              :string;
sTobRub             :string;
i                   :integer;
CheminSource        :string;
Buffer              :string;
F                   :TextFile;
Q1                  :TQuery;
DATEFIN             :TDateTime;
CODEACTIVITE        :string;
FRP,Command         :string;
Repsociete          :string;
St                  :string;
BEGIN
     Result := TRUE;
     if not FileExists(V_PGI_ENV.RepLocal+ '\APP\tdi_publifi.exe') then
     begin
          PGIINFO ('Attention Vous devrez installer TDI-Publifi' ,'TELETVA');
          RESULT := FALSE;
          exit;
     end;
     sSQL := 'SELECT * FROM '+LiasseEnv.FICHIERVALRUB;
     sSQL := sSQL + ' WHERE '+LiasseEnv.FVAL+'_MILLESIME = "'+LiasseEnv.MILLESIME+'"'+
          ' AND '+ LiasseEnv.FVAL+'_DATEEFFET = "'+USDATETIME(LiasseEnv.DATEENCOURS)+'"';

     QDos := OpenSQL(sSQL,true);
     if QDos.Eof then
     begin
          Ferme(Qdos);
          RESULT := FALSE;
          exit;
     end;
     if GetParamSocSecur ('SO_TELETVABQE1','') = '' then
     begin
          PGIINFO ('Il faut renseigner le compte de banque associé à la procédure TELETVA dans les paramètres sociétés ' ,'TELETVA');
          Ferme(Qdos);
          RESULT := FALSE;
          exit;
     end;
 //    CheminSource := V_PGI_ENV.RepLocal + '\STD\'+ VLIASSE^.RepDeclaration+ '\'+LiasseEnv.MILLESIME + '\Temp\Tva.txt';
     CheminSource := V_PGI_ENV.pathDos + '\Tva.txt';

     DeleteFile(Pchar(CheminSource));
 try
     AssignFile(F, CheminSource);	{ Ouverture du fichier de sortie }
     if FileExists (CheminSource) then  Append (F)
     else Rewrite (F);
     Q1 := OpenSql ('SELECT CA3_DATEFIN from CA3_DECLARATION Where CA3_STATUT="-"', TRUE);
     if not Q1.EOF then DATEFIN := Q1.FindField ('CA3_DATEFIN').Asdatetime;
     Ferme (Q1);
     Q1   := OpenSql('SELECT * FROM CA3_PARAMETRE', TRUE);
     if not Q1.EOF then
     begin
          CODEACTIVITE := Q1.FindField ('CP3_CODEACTIVE').Asstring;
          FRP := Q1.FindField ('CP3_RECETTE').Asstring + Q1.FindField ('CP3_DOSSIER').Asstring
          + Q1.FindField ('CP3_CLEDECLA').Asstring
     end;
     Ferme (Q1);

     Buffer :=  'DATE'+#9+FormatDateTime('yyyymmdd',LiasseEnv.DateEncours)
     +#9+FormatDateTime('yyyymmdd',DATEFIN);
     Writeln(F,Buffer);
     Buffer :=  '0011'+#9+FormatDateTime('yyyymmdd',LiasseEnv.DateEncours);
     Writeln(F,Buffer);
     Buffer := '0012'+#9+FormatDateTime('yyyymmdd',DATEFIN);
     Writeln(F,Buffer);
     Buffer := '0014'+#9+ V_PGI_ENV.NoDossier;
     Writeln(F,Buffer);

     TOB_TEMPO := TOB.Create('',Nil,-1);
     TOB_TEMPO.LoadDetailDB(LiasseEnv.FICHIERVALRUB,'','',QDos,true,true);
     Ferme(QDOS);
     // PARCOURS DE TOUTS LES RUB DE LA TOB
     //SI DE TYPE 0 (SAISIE) OU 2 (BD MODIFIE) ON ALIMENTE LE DICTIONAIRE DE DONNEE
     for i:= 0 to TOB_TEMPO.Detail.Count -1 do
     begin
          T := TOB_TEMPO.Detail[i];
          if T.GetValue(LiasseEnv.FVAL+'_NUMRUB')= '00950' then
             EcritureRib (F);
          if T.GetValue(LiasseEnv.FVAL+'_NUMRUB')= '00100' then
          begin
              if ctxPCL in V_PGI.PGIContexte then
              begin
                        Buffer := '00018'+#9+ 'CAB';
                        Writeln(F,Buffer);
              end
              else
              begin
                        Buffer := '00118'+#9+ 'ENT';
                        Writeln(F,Buffer);
              end;
          end;
          if T.GetValue(LiasseEnv.FVAL+'_NUMRUB')= '00400' then
          begin
                        Buffer := '00209'+#9+ CODEACTIVITE;
                        Writeln(F,Buffer);
                        Buffer := '00210'+#9+ GetParamSocSecur('SO_SIRET','');
                        Writeln(F,Buffer);
                        Buffer := '00211'+#9+ FRP;
                        Writeln(F,Buffer);
          end;

          Buffer := T.GetValue(LiasseEnv.FVAL+'_NUMRUB')+ #9 +
          T.GetValue(LiasseEnv.FVAL+'_VALRUB');
          Writeln(F,Buffer);
          if T.GetValue(LiasseEnv.FVAL+'_NUMRUB') = '00113' then
          begin
              Buffer := '00119'+#9+ GetParamSocSecur ('SO_CODEPOSTAL','');
              Writeln(F,Buffer);
              Buffer := '00120'+#9+ GetParamSocSecur ('SO_VILLE','');
              Writeln(F,Buffer);
          end ;
          if T.GetValue(LiasseEnv.FVAL+'_NUMRUB') = '00116' then
          begin
          if T.GetValue(LiasseEnv.FVAL+'_VALRUB') <> '' then
          begin
            Buffer := '00117'+#9+ Copy (T.GetValue(LiasseEnv.FVAL+'_VALRUB'), 0, 5);
            Writeln(F,Buffer);
            Buffer := '00118'+#9+ Copy (T.GetValue(LiasseEnv.FVAL+'_VALRUB'), 6, length (T.GetValue(LiasseEnv.FVAL+'_VALRUB'))-5);
            Writeln(F,Buffer);
          end;
          end ;
     end;

//     TOB_TEMPO.SaveToFile(sTempo,True,True,true);
  finally
   CloseFile(F);
  end;
     TOB_TEMPO.Free;
{
     St := V_PGI_ENV.PathDos;
     RepSociete := ReadTokenpipe (St, '\');
     RepSociete := RepSociete + '\'+ ReadTokenpipe (St, '\') + '\D000000\,';
}
     RepSociete := V_PGI_ENV.Com.dir+'\,';
     if not (ctxPCL in V_PGI.PGIContexte) then  RepSociete := 'PME';

  //   CheminSource := V_PGI_ENV.RepLocal + '\STD\'+ VLIASSE^.RepDeclaration+ '\'+LiasseEnv.MILLESIME + '\Temp\,';
      CheminSource := V_PGI_ENV.pathDos + '\,';

     Command := 'TVA,CA3,P,S,'+V_PGI_ENV.NoDossier+',TVA.txt,'+
     CheminSource + RepSociete + V_PGI_ENV.RepLocal + '\STD\';

     ShellExecute(0, PCHAR('open'),PCHAR('tdi_publifi.exe'),PCHAR(Command),Nil,SW_RESTORE);
   
END;

Procedure EcritureRib (var F :TextFile);
var
Q1                  : TQuery;
Buffer              : string;
begin
     // remplacement BQ_CODE  par BQ_GENERAL
       Q1 := OpenSql ('SELECT * from BANQUECP Where BQ_GENERAL="'+GetParamSocSecur ('SO_TELETVABQE1','')
                     +'" AND BQ_NODOSSIER = "'+V_PGI.NoDossier+'"', TRUE); // 19/10/2006 YMO Multisociétés
       if not Q1.EOF then
       begin
            if  Q1.FindField ('BQ_ETABBQ').asstring = '' then
            PGIINFO ('Il faut les RIB ' ,'TELETVA');

            Buffer := '00510'+#9+Q1.FindField ('BQ_ETABBQ').asstring;
            Writeln(F,Buffer);
            Buffer := '00513'+#9+Q1.FindField ('BQ_GUICHET').asstring;
            Writeln(F,Buffer);
            Buffer := '00514'+#9+Q1.FindField ('BQ_NUMEROCOMPTE').asstring+Q1.FindField ('BQ_CLERIB').asstring;
            Writeln(F,Buffer);
       end;
       Ferme (Q1);
       Q1 := OpenSql ('SELECT * from BANQUECP Where BQ_GENERAL="'+GetParamSocSecur ('SO_TELETVABQE2','')
                     +'" AND BQ_NODOSSIER = "'+V_PGI.NoDossier+'"', TRUE); // 19/10/2006 YMO Multisociétés
       if not Q1.EOF then
       begin
            if  Q1.FindField ('BQ_ETABBQ').asstring <> '' then
            begin
              Buffer := '00520'+#9+Q1.FindField ('BQ_ETABBQ').asstring;
              Writeln(F,Buffer);
              Buffer := '00523'+#9+Q1.FindField ('BQ_GUICHET').asstring;
              Writeln(F,Buffer);
              Buffer := '00524'+#9+Q1.FindField ('BQ_NUMEROCOMPTE').asstring+Q1.FindField ('BQ_CLERIB').asstring;
              Writeln(F,Buffer);
            end;
       end;
       Ferme (Q1);
       Q1 := OpenSql ('SELECT * from BANQUECP Where BQ_GENERAL="'+GetParamSocSecur ('SO_TELETVABQE3','')
                     +'" AND BQ_NODOSSIER = "'+V_PGI.NoDossier+'"', TRUE); // 19/10/2006 YMO Multisociétés
       if not Q1.EOF then
       begin
            if  Q1.FindField ('BQ_ETABBQ').asstring <> '' then
            begin
              Buffer := '00530'+#9+Q1.FindField ('BQ_ETABBQ').asstring;
              Writeln(F,Buffer);
              Buffer := '00533'+#9+Q1.FindField ('BQ_GUICHET').asstring;
              Writeln(F,Buffer);
              Buffer := '00534'+#9+Q1.FindField ('BQ_NUMEROCOMPTE').asstring+Q1.FindField ('BQ_CLERIB').asstring;
              Writeln(F,Buffer);
            end;
       end;
       Ferme (Q1);
end;
end.
