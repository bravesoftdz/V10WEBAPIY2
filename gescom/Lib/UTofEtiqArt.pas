unit UTofEtiqArt;

interface

uses  StdCtrls,Controls,Classes,forms,sysutils,ComCtrls,
      HCtrls,HEnt1,HMsgBox,UTOF, UTOB,
{$IFDEF EAGLCLIENT}
      utileAGL,
{$ELSE}
      db,{$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}Fiche,
{$IFDEF V530}
      EdtEtat,
{$ELSE}
      EdtREtat,
{$ENDIF}      
{$ENDIF}
      AglInit,EntGC, HQry ;

Procedure LanceEdtEtiquette(TOBL : TOB ; RegimePrix, CodeEtat : string);

Type
     TOF_EtiqArt = Class (TOF)
        procedure OnUpdate ; override ;
        procedure OnClose ; override ;
     end;


implementation

Procedure LanceEdtEtiquette(TOBL : TOB ; RegimePrix, CodeEtat : string);
var TobEtiq, TOBE : TOB;
    i_ind1, i_ind2 : integer;
begin
ExecuteSQL('DELETE FROM GCTMPETQ WHERE GZD_UTILISATEUR = "'+V_PGI.USer+'"');
TobEtiq := TOB.Create('GCTMPETQ',nil,-1);
for i_ind1:=0 to TOBL.Detail.Count-1 do
   for i_ind2:=0 to StrToInt(TOBL.Detail[i_ind1].GetValue('GL_QTEFACT'))-1 do
      begin
      TOBE:=TOB.Create('GCTMPETQ',TobEtiq,-1);
      TOBE.PutValue('GZD_UTILISATEUR',V_PGI.USer);
      TOBE.PutValue('GZD_COMPTEUR',i_ind2+1);
      TOBE.PutValue('GZD_ARTICLE',TOBL.Detail[i_ind1].GetValue('GL_ARTICLE'));
      TOBE.PutValue('GZD_REGIMEPRIX',RegimePrix);
      end;
if TobEtiq.Detail.Count > 0 then TobEtiq.InsertDB(nil);
LanceEtat( 'E','GEA',CodeEtat,True,False,False,nil,'GZD_UTILISATEUR="'+V_PGI.USer+'"','',False) ;
TobEtiq.Free;
AvertirTable('TTMODELETIQART') ;
ExecuteSQL('DELETE FROM GCTMPETQ WHERE GZD_UTILISATEUR = "'+V_PGI.USer+'"');
end;


procedure TOF_EtiqArt.OnUpdate ;
var
    stWhere, stInsert, RegimePrix : string;
    QEtiq : TQuery;
    i_ind1, nbEtiq, nbArticle : integer;
begin
Inherited;
ExecuteSQL('DELETE FROM GCTMPETQ WHERE GZD_UTILISATEUR = "'+V_PGI.USer+'"');
SetControlText('XX_WHERE_USER','');
stWhere := RecupWhereCritere(TPageControl(GetControl('PAGES')));
SetControlText('XX_WHERE_USER','GZD_UTILISATEUR="'+V_PGI.USer+'"');

QEtiq := OpenSQL('SELECT COUNT(GA_ARTICLE) FROM ARTICLE ' + stWhere, true);
nbArticle := QEtiq.Fields[0].AsInteger;
Ferme(QEtiq);
if nbArticle = 0 then exit;

if (ctxMode in V_PGI.PGIContexte) then
    RegimePrix := 'TTC'
    else
    if GetControlText('RHT') = 'X' then RegimePrix := 'HT' else RegimePrix := 'TTC';

nbEtiq := StrToInt(GetControlText('NBEXEMPLAIRE'));
ExecuteSQL('Insert into GCTMPETQ (GZD_UTILISATEUR, GZD_COMPTEUR, GZD_ARTICLE) values ("' +
            V_PGI.User + '", 0, "XXX")');
stInsert := 'Insert into GCTMPETQ (GZD_UTILISATEUR, GZD_COMPTEUR, GZD_ARTICLE, GZD_REGIMEPRIX) ' +
            'Select "' + V_PGI.User + '" as GZD_UTILISATEUR, ' +
            '(Select Max(GZD_COMPTEUR) from GCTMPETQ where GZD_UTILISATEUR="'+ V_PGI.User +'") + 1 as GZD_COMPTEUR, ' +
            'GA_ARTICLE as GZD_ARTICLE, ' +
            '"' + RegimePrix + '" as GZD_REGIMEPRIX from ARTICLE';
for i_ind1:=0 to nbEtiq - 1 do ExecuteSQL(stInsert);
ExecuteSQL('Delete from GCTMPETQ where GZD_UTILISATEUR="'+ V_PGI.User +'" and GZD_COMPTEUR=0 and GZD_ARTICLE="XXX"');
end;

procedure TOF_EtiqArt.OnClose ;
begin
Inherited;
ExecuteSQL('DELETE FROM GCTMPETQ WHERE GZD_UTILISATEUR = "'+V_PGI.USer+'"');
end;


Initialization
registerclasses([TOF_EtiqArt]);

end.
