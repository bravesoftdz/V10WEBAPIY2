unit UTofEtiqArtStk;

interface

uses  StdCtrls,Controls,Classes,forms,sysutils,ComCtrls,
      HCtrls,HEnt1,HMsgBox,UTOF, UTOB,CalcOleGescom,
{$IFDEF EAGLCLIENT}
      eQRS1,utileAGL,
{$ELSE}
      QRS1,db,dbTables,EdtEtat,Fiche,
{$ENDIF}
      AglInit,EntGC,HQry,UtilArticle,UtilGc ;


Type
     TOF_EtiqArtStk = Class (TOF)
        procedure OnUpdate ; override ;
        procedure OnClose ; override ;
        procedure OnArgument (Argument : String ) ; override ;
     end;

     var Critere : string;

implementation
procedure TOF_EtiqArtStk.OnArgument (Argument : String ) ;
var stArgument : string;
    iFam : integer ;
begin
inherited ;
stArgument := Argument;
Critere:=uppercase(Trim(ReadTokenSt(stArgument))) ;

// Paramétrage des libellés des familles
for iFam:=1 to 3 do ChangeLibre2('TGA_FAMILLENIV'+IntToStr(iFam),Ecran);
if (ctxMode in V_PGI.PGIContexte) and (GetPresentation=ART_ORLI) then
    for iFam:=4 to 8 do ChangeLibre2('TGA2_FAMILLENIV'+IntToStr(iFam),Ecran);

end;


procedure TOF_EtiqArtStk.OnUpdate ;
var F : TFQRS1;
    stWhere, RegimePrix, Sql: string;
    QEtiq : TQuery;
    TobTemp, TobEtiq, TOBE : TOB;
    i_ind1, i_ind2, nbex : integer;
    Edit2 : THValcombobox ;
    ctrl, ctrlName,signe, NBEXEMPLAIRE,CodeArticleGen : string ;
    CodeEtat, LibEtat : string;
begin
Inherited;
VH_GC.TOBEdt.ClearDetail ;
initvariable;
F := TFQRS1(Ecran);
ExecuteSQL('DELETE FROM GCTMPETQ WHERE GZD_UTILISATEUR = "'+V_PGI.USer+'"');
SetControlText('XX_WHERE_USER','');
// SetControlText('XX_WHERE','');  OT mis en commentaire car pas utilisé  FQ 10829
//récupération du code et du libellé de l'état
CodeEtat := GetControlText ('FETAT');
LibEtat := RechDom('TTMODELETIQART',CodeEtat,FALSE);

stWhere := RecupWhereCritere(TPageControl(F.FindComponent('PAGES')));
if stWhere<>'' then stWhere:=stWhere+' AND '
else stWhere:=' WHERE ' ;

// Prise en compte uniquement du stock non clôturé.
stWhere:=stWhere+'GQ_CLOTURE="-" AND GQ_PHYSIQUE>0 ' ;

ctrl:='GQ_DEPOT';
ctrlName:='ETABLISSEMENT'; signe:='=' ;
Edit2:=THValComboBox(TFQRS1(F).FindComponent(ctrlName)) ;
if (Edit2 <> nil) and (Edit2.Value <> '') and (Edit2.Value <> TraduireMemoire ('<<Tous>>'))
  then stWhere:=stWhere+' AND '+ctrl+signe+'"'+Edit2.Value+'"' ;

SetControlText('XX_WHERE_USER','GZD_UTILISATEUR="'+V_PGI.USer+'"');
if (ctxMode in V_PGI.PGIContexte) and (GetPresentation=ART_ORLI) then
    Sql:= 'SELECT GQ_DEPOT,GQ_ARTICLE,GA_CODEARTICLE,GQ_PHYSIQUE FROM DISPODIM_MODES5 '+ stWhere
else
    Sql:= 'SELECT GQ_DEPOT,GQ_ARTICLE,GA_CODEARTICLE,GQ_PHYSIQUE FROM DISPODIM_MODE '+ stWhere ;

RegimePrix := 'TTC';
QEtiq:=OpenSQL(Sql,False) ;
//if QEtiq.Eof then begin Ferme(QEtiq); exit; end;
TobTemp:=TOB.Create('',nil,-1);
if not QEtiq.Eof then TobTemp.LoadDetailDB('','','',QEtiq,false)
else
   begin
   TobTemp.free;
   TobTemp:=nil;
   end;
Ferme(QEtiq);
if TobTemp<>nil then
   begin
   TobEtiq := TOB.Create('GCTMPETQ',nil,-1);
   if EtatMonarchFactorise(LibEtat) then
      begin
      for i_ind1:=0 to TobTemp.Detail.Count-1 do
         begin
         TOBE:=TOB.Create('GCTMPETQ',TobEtiq,-1);
         TOBE.PutValue('GZD_UTILISATEUR',V_PGI.USer);
         TOBE.PutValue('GZD_COMPTEUR',i_ind1);
         TOBE.PutValue('GZD_ARTICLE',TobTemp.Detail[i_ind1].GetValue('GQ_ARTICLE'));
         TOBE.PutValue('GZD_CODEARTICLE',TobTemp.Detail[i_ind1].GetValue('GA_CODEARTICLE'));
         TOBE.PutValue('GZD_DEPOT',TobTemp.Detail[i_ind1].GetValue('GQ_DEPOT'));
         TOBE.PutValue('GZD_REGIMEPRIX',RegimePrix);
         TOBE.PutValue('GZD_NBETIQDIM',TobTemp.Detail[i_ind1].GetValue('GQ_PHYSIQUE'));
         // Appel de la fonction qui transforme le codearticle en article générique (avec le X à la fin)
         CodeArticleGen:=GCDimToGen(TobTemp.Detail[i_ind1].GetValue('GA_CODEARTICLE'));
         TOBE.PutValue('GZD_CODEARTICLEGEN',CodeArticleGen);
         end;
      end
   else
      begin
      //i_ind2:=0;
      for i_ind1:=0 to TobTemp.Detail.Count-1 do
         begin
         NBEXEMPLAIRE := TobTemp.Detail[i_ind1].GetValue('GQ_PHYSIQUE');
         nbex := StrToInt(NBEXEMPLAIRE);
         for i_ind2:=1 to nbex do
            begin
            TOBE:=TOB.Create('GCTMPETQ',TobEtiq,-1);
            TOBE.PutValue('GZD_UTILISATEUR',V_PGI.USer);
            TOBE.PutValue('GZD_COMPTEUR',i_ind2);
            TOBE.PutValue('GZD_ARTICLE',TobTemp.Detail[i_ind1].GetValue('GQ_ARTICLE'));
            TOBE.PutValue('GZD_CODEARTICLE',TobTemp.Detail[i_ind1].GetValue('GA_CODEARTICLE'));
            TOBE.PutValue('GZD_DEPOT',TobTemp.Detail[i_ind1].GetValue('GQ_DEPOT'));
            TOBE.PutValue('GZD_REGIMEPRIX',RegimePrix);
            // Appel de la fonction qui transforme le codearticle en article générique (avec le X à la fin)
            CodeArticleGen:=GCDimToGen(TobTemp.Detail[i_ind1].GetValue('GA_CODEARTICLE'));
            TOBE.PutValue('GZD_CODEARTICLEGEN',CodeArticleGen);
            end;
         end;
      end;
   if TobEtiq.Detail.Count > 0 then TobEtiq.InsertDB(nil); 
   TobTemp.Free;
   TobEtiq.Free;
   end;
EditMonarchSiEtat (LibEtat);
end;

procedure TOF_EtiqArtStk.OnClose ;
begin
Inherited;
EditMonarch ('');
ExecuteSQL('DELETE FROM GCTMPETQ WHERE GZD_UTILISATEUR = "'+V_PGI.USer+'"');
VH_GC.TOBEdt.ClearDetail ;
initvariable;
end;


Initialization
registerclasses([TOF_EtiqArtStk]);

end.
