unit UTofEdEtVte;

interface

uses  StdCtrls,Controls,Classes,forms,sysutils,ComCtrls,
      HCtrls,HEnt1,HMsgBox,UTOF, UTOB,Spin,UtilSynVte,
{$IFDEF EAGLCLIENT}
      eQRS1,utileAGL,
{$ELSE}
      QRS1,db,dbTables,Fiche,
{$IFDEF V530}
      EdtEtat,
{$ELSE}
      EdtREtat,
{$ENDIF}      
{$ENDIF}
      AglInit,EntGC, HQry,HStatus ,MajTable,UtilArticle,UtilGC;

Type
     TOF_EdEtVte = Class (TOF)
        procedure OnLoad  ; override ;
        procedure OnUpdate ; override ;
        procedure OnClose ; override ;
        procedure OnArgument (Argument : String ) ; override ;
     private
        QteVtenul, QteStonul, QteAttnul : boolean;
        datdeb, datfin : TDateTime ;
     end;

implementation
procedure TOF_EdEtVte.OnArgument (Argument : String ) ;
var stArgument : string;
    iCol : integer;
begin
inherited ;
stArgument := Argument;
// Si le libellé de la famille est ".-" alors on ne l'affiche pas
// Je ne peut pas utiliser ChangeLibre2 car je ne veux pas mettre les critères dans la
// clause where de l'état donc je ne dois pas nommé mes champs avec "GA_" or ChangeLibre2 ne
// fonctionne que comme ça.
if (ctxMode in V_PGI.PGIContexte) and (GetPresentation=ART_ORLI) then
    begin
    for iCol:=4 to 8 do
       begin
       if THLabel(GetControl('T_FAMILLENIV'+InttoStr(iCol))).Caption='.-' then
          begin
          SetControlVisible('T_FAMILLENIV'+InttoStr(iCol),False);
          SetControlVisible('FAMILLENIV'+InttoStr(iCol),False);
          end;
       end;
    end
else if (ctxMode in V_PGI.PGIContexte) then // Si on n'est pas en presentation orli on n'affiche que 3 niveau de famille
   begin
   for iCol:=4 to 8 do
      begin
      setControlVisible('T_FAMILLENIV'+InttoStr(iCol),GetPresentation=ART_ORLI);
      setControlVisible('FAMILLENIV'+InttoStr(iCol),GetPresentation=ART_ORLI);
      end;
   end;

// Si on est pas en mode ORLI, on ne peut pas regrouper et trier dans les 8 niveaux de famille
if (ctxMode in V_PGI.PGIContexte) and not (GetPresentation=ART_ORLI) then
    begin
    THValComboBox(GetControl('REGROUP')).Plus := ' AND CO_CODE not in ("NV4","NV5","NV6","NV7","NV8") ';
    THValComboBox(GetControl('ORDER')).Plus := ' AND CO_CODE not in ("NV4","NV5","NV6","NV7","NV8") ';
    end;

THValComboBox(GetControl('REGROUP')).Value := 'ART';
end;

procedure TOF_EdEtVte.OnLoad  ;
begin
inherited ;
SetControlText('XX_WHERE_USER','');
SetControlText('XX_WHERE_USER','GZL_UTILISATEUR="'+V_PGI.USer+'"');
end ;

function FindStInTAB(St:String;Tab:Array of string):boolean;
var i : integer;
begin
result:=false;
for i:=Low(Tab) to High(Tab) do
  if Pos(St,Tab[i])<>0 then begin result:=true; break; end;
end;

procedure TOF_EdEtVte.OnUpdate ;
var F : TFQRS1;
    stWhere, ston, SqlStock, SqlAttendu, SqlVente, StRupt,stFrom : string;
    QLigne,QQ : TQuery;
    TobStock,TobVente,TobAttendu, TobT  : TOB;
    Edit : THEdit ;
    Edit2 : THMultiValcombobox ;
    Edit3 : THValcombobox ;
    Edit4 : TSpinEdit ;
    nbmois,i_ind,CountStArt,NbRequeteArt,y,MoisVente,AnneeVente,noVente,iCol : integer ;
    ctrl,ctrlLigne,ctrlName,signe,tri,rupt,sttri,rupt0,ListeCritere: string ;
    codeart_sv, codedep_sv, StArticle, StDepot, stJoin, librupt, libtri, prefixeRupt, prefixeTri : string ;
    qte_stock,qte_attendu : double;
    TabWhere,TabWhereDepot : array of String;
    Annee, Mois, Jour : Word;
begin
Inherited;
F := TFQRS1(Ecran);
QteVtenul := False;
QteStonul := False;
QteAttnul := False;
if TcheckBox(TFQRS1(F).FindComponent('QTEVTENUL')).State = cbChecked then
   QteVtenul := True;
if TcheckBox(TFQRS1(F).FindComponent('QTESTONUL')).State = cbChecked then
   QteStonul := True;
if TcheckBox(TFQRS1(F).FindComponent('QTEATTNUL')).State = cbChecked then
   QteAttnul := True;

ExecuteSQL('DELETE FROM GCTMPETATVTE WHERE GZL_UTILISATEUR = "'+V_PGI.USer+'"');

ctrl:='GQ_DEPOT';
ctrlLigne:='GL_ETABLISSEMENT';
ctrlName:='ETABLISSEMENT'; signe:=' in ' ;
Edit2:=THMultiValComboBox(TFQRS1(F).FindComponent(ctrlName)) ;
if (Edit2<>nil) and (Edit2.Text<>'') and (Edit2.Text <> TraduireMemoire ('<<Tous>>')) then
  begin
  // Pour que la liste des établissements soit cohérent avec une clause IN je remplace les ';' en '","'
  ListeCritere:=StringReplace(copy(Edit2.Text,1,length(Edit2.Text)-1),';','","',[rfReplaceAll]);
  // Clause where pour se qui concerne le stock
  if stWhere<>'' then stWhere:=stWhere+' AND ' ;
  stWhere:=stWhere+ctrl+signe+'("'+ListeCritere+'")' ;
  // Clause where pour ce qui concerne l'attendu et les ventes
  if ston<>'' then ston:=ston+' AND ' ;
  ston:=ston+ctrlLigne+signe+'("'+ListeCritere+'")' ;
  end;

ctrl:='GA_CODEARTICLE';
ctrlLigne:= 'GL_CODEARTICLE';
ctrlName:='CODEARTICLE'; signe:='=' ;
Edit:=THEdit(TFQRS1(F).FindComponent(ctrlName)) ;
if (Edit<>nil) and (Edit.Text<>'') then
  begin
  // Clause where pour se qui concerne le stock
  if stWhere<>'' then stWhere:=stWhere+' AND ' ;
  stWhere:=stWhere+ctrl+signe+'"'+Edit.Text+'"' ;
  // Clause where pour ce qui concerne l'attendu et les ventes
  if ston<>'' then ston:=ston+' AND ' ;
  ston:=ston+ctrlLigne+signe+'"'+Edit.Text+'"' ;
  end ;

ctrl:='GA_FOURNPRINC';
ctrlLigne:= 'GL_TIERS';
ctrlName:='FOURNPRINC'; signe:='=' ;
Edit:=THEdit(TFQRS1(F).FindComponent(ctrlName)) ;
if (Edit<>nil) and (Edit.Text<>'') then
  begin
  // Clause where pour se qui concerne le stock
  if stWhere<>'' then stWhere:=stWhere+' AND ' ;
  stWhere:=stWhere+ctrl+signe+'"'+Edit.Text+'"' ;
  // Clause where pour ce qui concerne l'attendu et les ventes
  if ston<>'' then ston:=ston+' AND ' ;
  ston:=ston+ctrlLigne+signe+'"'+Edit.Text+'"' ;
  end ;

ctrl:='GA_COLLECTION';
ctrlLigne:= 'GL_COLLECTION';
ctrlName:='COLLECTION'; signe:=' in ' ;
Edit2:=THMultiValComboBox(TFQRS1(F).FindComponent(ctrlName)) ;
if (Edit2<>nil) and (Edit2.Text<>'') and (Edit2.Text <> TraduireMemoire ('<<Tous>>')) then
  begin
  // Pour que la liste des établissements soit cohérent avec une clause IN je remplace les ';' en '","'
  ListeCritere:=StringReplace(copy(Edit2.Text,1,length(Edit2.Text)-1),';','","',[rfReplaceAll]);
  // Clause where pour se qui concerne le stock
  if stWhere<>'' then stWhere:=stWhere+' AND ' ;
  stWhere:=stWhere+ctrl+signe+'("'+ListeCritere+'")' ;
  // Clause where pour ce qui concerne l'attendu et les ventes
  if ston<>'' then ston:=ston+' AND ' ;
  ston:=ston+ctrlLigne+signe+'("'+ListeCritere+'")' ;
  end ;

// Récupération des 3 niveaux de familles
for iCol:=1 to 3 do
   begin
   ctrl:='GA_FAMILLENIV'+IntToStr(iCol);
   ctrlLigne:= 'GL_FAMILLENIV'+IntToStr(iCol);
   ctrlName:='FAMILLENIV'+IntToStr(iCol); signe:=' in ' ;
   Edit2:=THMultiValComboBox(TFQRS1(F).FindComponent(ctrlName)) ;
   if (Edit2<>nil) and (Edit2.Text<>'') and (Edit2.Text <> TraduireMemoire ('<<Tous>>')) then
     begin
     // Pour que la liste des établissements soit cohérent avec une clause IN je remplace les ';' en '","'
     ListeCritere:=StringReplace(copy(Edit2.Text,1,length(Edit2.Text)-1),';','","',[rfReplaceAll]);
     // Clause where pour se qui concerne le stock
     if stWhere<>'' then stWhere:=stWhere+' AND ' ;
     stWhere:=stWhere+ctrl+signe+'("'+ListeCritere+'")' ;
     // Clause where pour ce qui concerne l'attendu et les ventes
     if ston<>'' then ston:=ston+' AND ' ;
     ston:=ston+ctrlLigne+signe+'("'+ListeCritere+'")' ;
     end ;
   end;

// Si on est en mode ORLI, je récupère les 5 autres niveaux de familles
if (ctxMode in V_PGI.PGIContexte) and (GetPresentation=ART_ORLI) then
   begin
   for iCol:=4 to 8 do
      begin
      ctrl:='GA2_FAMILLENIV'+IntToStr(iCol);
      ctrlLigne:= ctrl;
      ctrlName:='FAMILLENIV'+IntToStr(iCol); signe:=' in ' ;
      Edit2:=THMultiValComboBox(TFQRS1(F).FindComponent(ctrlName)) ;
      if (Edit2<>nil) and (Edit2.Text<>'') and (Edit2.Text <> TraduireMemoire ('<<Tous>>')) then
         begin
         // Pour que la liste des établissements soit cohérent avec une clause IN je remplace les ';' en '","'
         ListeCritere:=StringReplace(copy(Edit2.Text,1,length(Edit2.Text)-1),';','","',[rfReplaceAll]);
         // Clause where pour se qui concerne le stock
         if stWhere<>'' then stWhere:=stWhere+' AND ' ;
         stWhere:=stWhere+ctrl+signe+'("'+ListeCritere+'")' ;
         // Clause where pour ce qui concerne l'attendu et les ventes
         if ston<>'' then ston:=ston+' AND ' ;
         ston:=ston+ctrlLigne+signe+'("'+ListeCritere+'")' ;
         end ;
      end;
   end;

ctrl:='GL_DATEPIECE';
Edit:=THEdit(TFQRS1(F).FindComponent('DATJOUR')) ;
datfin := StrToDate(Edit.Text);
datfin := FinDeMois(datfin);
edit4 := TSpinEdit(TFQRS1(F).FindComponent('NBMOIS'));
nbmois := edit4.value;
datdeb:=PlusMois(datfin,-nbmois) ;
datdeb := DebutDeMois(datdeb);
if TcheckBox(TFQRS1(F).FindComponent('ENCOURS')).State <> cbChecked then
       datfin:=PlusMois(datfin,-1) ;
if ston<>'' then ston:=ston+' AND ' ;
ston:=ston + ctrl+ ' <= "'+USDateTime(datfin)+'"' ;

Edit3:=THValComboBox(GetControl('REGROUP'));

QQ:=OpenSQL('SELECT CO_ABREGE FROM COMMUN WHERE CO_TYPE="GVE" AND CO_CODE="'+ Edit3.Value +'"', True);
rupt := QQ.Fields[0].AsString;
stJoin :='';
if rupt = 'GA_CODEARTICLE' then
   begin
   prefixeRupt:='GL_';
   librupt :='GA_LIBELLE'
   end
else if rupt = 'GA_FAMILLENIV1' then
     begin
     prefixeRupt:='GL_';
     librupt := 'CC1R.CC_LIBELLE';
     stJoin := 'LEFT OUTER JOIN CHOIXCOD CC1R ON CC1R.CC_CODE=GA_FAMILLENIV1 AND CC1R.CC_TYPE="FN1" ';
     end
else if rupt = 'GA_FAMILLENIV2' then
     begin
     prefixeRupt:='GL_';
     librupt := 'CC2R.CC_LIBELLE';
     stJoin := 'LEFT OUTER JOIN CHOIXCOD CC2R ON CC2R.CC_CODE=GA_FAMILLENIV2 AND CC2R.CC_TYPE="FN2" ';
     end
else if rupt = 'GA_FAMILLENIV3' then
     begin
     prefixeRupt:='GL_';
     librupt :='CC3R.CC_LIBELLE';
     stJoin := 'LEFT OUTER JOIN CHOIXCOD CC3R ON CC3R.CC_CODE=GA_FAMILLENIV3 AND CC3R.CC_TYPE="FN3" ';
     end
else if rupt = 'GA2_FAMILLENIV4' then
     begin
     prefixeRupt:='GA2_'; // Ici le prefixe est différent car dans la table ligne il n'y a pas les niveaux de famille > 8
     librupt :='CC4R.CC_LIBELLE';
     stJoin := 'LEFT OUTER JOIN CHOIXCOD CC4R ON CC4R.CC_CODE=GA2_FAMILLENIV4 AND CC4R.CC_TYPE="FN4" ';
     end
else if rupt = 'GA2_FAMILLENIV5' then
     begin
     prefixeRupt:='GA2_';
     librupt :='CC5R.CC_LIBELLE';
     stJoin := 'LEFT OUTER JOIN CHOIXCOD CC5R ON CC5R.CC_CODE=GA2_FAMILLENIV5 AND CC5R.CC_TYPE="FN5" ';
     end
else if rupt = 'GA2_FAMILLENIV6' then
     begin
     prefixeRupt:='GA2_';
     librupt :='CC6R.CC_LIBELLE';
     stJoin := 'LEFT OUTER JOIN CHOIXCOD CC6R ON CC6R.CC_CODE=GA2_FAMILLENIV6 AND CC6R.CC_TYPE="FN6" ';
     end
else if rupt = 'GA2_FAMILLENIV7' then
     begin
     prefixeRupt:='GA2_';
     librupt :='CC7R.CC_LIBELLE';
     stJoin := 'LEFT OUTER JOIN CHOIXCOD CC7R ON CC7R.CC_CODE=GA2_FAMILLENIV7 AND CC7R.CC_TYPE="FN7" ';
     end
else if rupt = 'GA2_FAMILLENIV8' then
     begin
     prefixeRupt:='GA2_';
     librupt :='CC8R.CC_LIBELLE';
     stJoin := 'LEFT OUTER JOIN CHOIXCOD CC8R ON GA2_FAMILLENIV8=CC8R.CC_CODE AND CC8R.CC_TYPE="FN8" ';
     end
else if rupt = 'GA_COLLECTION' then
     begin
     prefixeRupt:='GL_';
     librupt := 'CC9R.CC_LIBELLE';
     stJoin := 'LEFT OUTER JOIN CHOIXCOD CC9R ON CC9R.CC_CODE=GA_COLLECTION AND CC9R.CC_TYPE="GCO" ';
     end
else if rupt = 'GA_FOURNPRINC' then
     begin
     prefixeRupt:='GL_';
     librupt := 'T_LIBELLE';
     stJoin := 'LEFT OUTER JOIN TIERS ON GA_FOURNPRINC=T_TIERS ';
     end
else if rupt = 'GQ_DEPOT' then
     begin
     prefixeRupt:='GL_';
     librupt := 'ET_LIBELLE';
     stJoin := 'LEFT OUTER JOIN ETABLISS ON GQ_DEPOT=ET_ETABLISSEMENT ';
     end;

Edit3:=THValComboBox(GetControl('ORDER'));

QQ:=OpenSQL('SELECT CO_ABREGE FROM COMMUN WHERE CO_TYPE="GVE" AND CO_CODE="'+ Edit3.Value +'"', True);
tri := QQ.Fields[0].AsString;
if tri = 'GA_FAMILLENIV1' then
     begin
     prefixeTri:='GL_';
     libtri := 'CC1T.CC_LIBELLE';
     stJoin := stJoin + 'LEFT OUTER JOIN CHOIXCOD CC1T ON CC1T.CC_CODE=GA_FAMILLENIV1 AND CC1T.CC_TYPE="FN1" ';
     end
else if tri = 'GA_FAMILLENIV2' then
     begin
     prefixeTri:='GL_';
     libtri := 'CC2T.CC_LIBELLE';
     stJoin := stJoin + 'LEFT OUTER JOIN CHOIXCOD CC2T ON CC2T.CC_CODE=GA_FAMILLENIV2 AND CC2T.CC_TYPE="FN2" ';
     end
else if tri = 'GA_FAMILLENIV3' then
     begin
     prefixeTri:='GL_';
     libtri :='CC3T.CC_LIBELLE';
     stJoin := stJoin + 'LEFT OUTER JOIN CHOIXCOD CC3T ON CC3T.CC_CODE=GA_FAMILLENIV3 AND CC3T.CC_TYPE="FN3" ';
     end
else if tri = 'GA2_FAMILLENIV4' then
     begin
     prefixeTri:='GA2_';
     libtri :='CC4T.CC_LIBELLE';
     stJoin := stJoin + 'LEFT OUTER JOIN CHOIXCOD CC4T ON CC4T.CC_CODE=GA2_FAMILLENIV4 AND CC4T.CC_TYPE="FN4" ';
     end
else if tri = 'GA2_FAMILLENIV5' then
     begin
     prefixeTri:='GA2_';
     libtri :='CC5T.CC_LIBELLE';
     stJoin := stJoin + 'LEFT OUTER JOIN CHOIXCOD CC5T ON CC5T.CC_CODE=GA2_FAMILLENIV5 AND CC5T.CC_TYPE="FN5" ';
     end
else if tri = 'GA2_FAMILLENIV6' then
     begin
     prefixeTri:='GA2_';
     libtri :='CC6T.CC_LIBELLE';
     stJoin := stJoin + 'LEFT OUTER JOIN CHOIXCOD CC6T ON CC6T.CC_CODE=GA2_FAMILLENIV6 AND CC6T.CC_TYPE="FN6" ';
     end
else if tri = 'GA2_FAMILLENIV7' then
     begin
     prefixeTri:='GA2_';
     libtri :='CC7T.CC_LIBELLE';
     stJoin := stJoin + 'LEFT OUTER JOIN CHOIXCOD CC7T ON CC7T.CC_CODE=GA2_FAMILLENIV7 AND CC7T.CC_TYPE="FN7" ';
     end
else if tri = 'GA2_FAMILLENIV8' then
     begin
     prefixeTri:='GA2_';
     libtri :='CC8T.CC_LIBELLE';
     stJoin := stJoin + 'LEFT OUTER JOIN CHOIXCOD CC8T ON CC8T.CC_CODE=GA2_FAMILLENIV8 AND CC8T.CC_TYPE="FN8" ';
     end
else if tri = 'GA_COLLECTION' then
     begin
     prefixeTri:='GL_';
     libtri := 'CC9T.CC_LIBELLE';
     stJoin := stJoin + 'LEFT OUTER JOIN CHOIXCOD CC9T ON CC9T.CC_CODE=GA_COLLECTION AND CC9T.CC_TYPE="GCO" ';
     end
else if tri = 'GA_FOURNPRINC' then
     begin
     prefixeTri:='GL_';
     libtri := 'T_LIBELLE';
     stJoin := stJoin + 'LEFT OUTER JOIN TIERS ON GA_FOURNPRINC=T_TIERS ';
     end
else if tri = 'GQ_DEPOT' then
     begin
     prefixeTri:='GL_';
     libtri := 'ET_LIBELLE';
     stJoin := stJoin + 'LEFT OUTER JOIN ETABLISS ON GQ_DEPOT=ET_ETABLISSEMENT ';
     end;

// 1. Chargement des articles et qtés en stock
// ceci est fait en fonction de la rupture et du tri demandé
// ===========================================
if (ctxMode in V_PGI.PGIContexte) and (GetPresentation=ART_ORLI) then stFrom:='from DISPOART_MODES5 '
else stFrom:='from DISPOART_MODE ';

SqlStock := 'SELECT '+ rupt + ' as GZL_RUPT1, max('+librupt+') as GZL_LIBRUPT1, ';
if tri<>'' then SqlStock := SqlStock + tri + ' as GZL_TRI, max('+libtri+') as GZL_LIBTRI, ';
SqlStock := SqlStock + 'sum(GQ_PHYSIQUE) as GZL_PHYSIQUE, "' + V_PGI.USer + '" as GZL_UTILISATEUR '
+stFrom + stjoin + ' where GA_TYPEARTICLE = "MAR" and GQ_CLOTURE<>"X" ' ;
if stwhere<>'' then SqlStock := SqlStock + 'AND ' + stwhere ;
SqlStock := SqlStock + ' group by ' + rupt ;
if tri<>'' then
   begin
   if (tri<>rupt) then SqlStock := SqlStock + ',' + tri;
   SqlStock := SqlStock + ' order by GZL_TRI ';
   end;

QLigne := OpenSQL(SqlStock, True);

if QLigne.Eof then
   begin
   Ferme(QLigne);
   exit;
   end;

TobStock:=TOB.Create('GCTMPETATVTE_',nil,-1);
TobStock.LoadDetailDB('GCTMPETATVTE','','',QLigne,false);
Ferme(QLigne);
if rupt='GA_FOURNPRINC' then rupt:='TIERS'
else rupt:= copy(rupt,pos('_',rupt)+1,length(rupt));

if tri='GA_FOURNPRINC' then tri:='TIERS'
else tri:= copy(tri,pos('_',tri)+1,length(tri));

// Recherche des quantités attendues dans la table ligne et je les insere dans TobStock
SqlAttendu:='SELECT '+ prefixeRupt + rupt + ' as RUPT, ';
if tri<>'' then SqlAttendu:=SqlAttendu + prefixeTri + tri + ' as TRI, ';
SqlAttendu:=SqlAttendu + 'sum (GL_QTEFACT) as ATTENDU FROM LIGNE ';
if (prefixeRupt='GA2_') or (prefixeTri='GA2_') then SqlAttendu:=SqlAttendu + 'left join articlecompl on GA2_CODEARTICLE=GL_CODEARTICLE ';
SqlAttendu:=SqlAttendu + 'WHERE GL_NATUREPIECEG = "CF" ';
if ston<>'' then SqlAttendu := SqlAttendu + 'AND ' + ston ;
SqlAttendu := SqlAttendu + ' AND GL_VIVANTE = "X" AND GL_CODEARTICLE<>"" '+
   ' GROUP BY '+prefixeRupt+rupt;
if (tri<>'') and (tri<>rupt) then SqlAttendu:=SqlAttendu + ', '+ prefixeTri + tri ;
QLigne := OpenSQL(SqlAttendu, True);

if not QLigne.Eof then
   begin
   TobAttendu:=TOB.Create('',nil,-1);
   TobAttendu.LoadDetailDB('','','',QLigne,false);
   Ferme(QLigne);
   for i_ind:=0 to TobAttendu.Detail.count-1 do
      begin
           if not VarIsNull(TobAttendu.Detail[i_ind].GetValue('RUPT')) then
              StRupt :=  TobAttendu.Detail[i_ind].GetValue('RUPT')
           else StRupt:='';
           qte_attendu := TobAttendu.Detail[i_ind].GetValue('ATTENDU');
           if tri<>'' then
              begin
              if not VarIsNull(TobAttendu.Detail[i_ind].GetValue('TRI')) then
                 StTri := TobAttendu.Detail[i_ind].GetValue('TRI')
              else StTri := '';
              TobT := TobStock.FindFirst(['GZL_RUPT1','GZL_TRI'],[StRupt,StTri],True);
              end
           else TobT := TobStock.FindFirst(['GZL_RUPT1'],[StRupt],True);
           if TobT<>nil then  TobT.PutValue('GZL_ATTENDU',qte_attendu );
      end;
   TobAttendu.free;
   end
else Ferme(QLigne);

// 2. Chargement des qtés vendues par période
// ==========================================
ston:=ston+' AND '+ctrl+' >= "'+USDateTime(datdeb)+'"' ;

SqlVente:='SELECT '+prefixeRupt+rupt+' as RUPT, ';
if tri<>'' then
   begin
   if tri='DEPOT' then tri:='ETABLISSEMENT'
   else if tri='TIERS' then tri:='FOURNISSEUR';
   SqlVente:=SqlVente + prefixeTri + tri + ' as TRI, ';
   end;
SqlVente:=SqlVente + 'sum(GL_QTEFACT) as QTEVTE, '+DB_Month('GL_DATEPIECE')+' as MOIS, '+
DB_YEAR('GL_DATEPIECE')+' as ANNEE '+ 'FROM LIGNE ';
if (prefixeRupt='GA2_') or (prefixeTri='GA2_') then SqlVente:=SqlVente + 'left join articlecompl on GA2_CODEARTICLE=GL_CODEARTICLE ';
SqlVente:=SqlVente + 'WHERE GL_TYPEARTICLE = "MAR" AND GL_NATUREPIECEG="FFO" ';
if ston<>'' then SqlVente := SqlVente + 'AND ' + ston ;
SqlVente := SqlVente + ' group BY '+prefixeRupt+rupt+', '+DB_MONTH('GL_DATEPIECE')+', '+DB_YEAR('GL_DATEPIECE');
if (tri<>'') and (tri<>rupt) then SqlVente:=SqlVente + ', '+ prefixeTri + tri ;
SqlVente := SqlVente +' order BY '+prefixeRupt+rupt;
if (tri<>'') and (tri<>rupt) then SqlVente:=SqlVente + ', '+ prefixeTri + tri ;
SqlVente := SqlVente +', '+DB_MONTH('GL_DATEPIECE')+', '+DB_YEAR('GL_DATEPIECE');

QLigne := OpenSQL(SqlVente, True);

if not QLigne.Eof then
   begin
   TobVente:=TOB.Create('',nil,-1);
   TobVente.LoadDetailDB('','','',QLigne,false);
   Ferme(QLigne);
   // Récupération de l'année, le mois et le jour de la date de fin
   DecodeDate(datFin, Annee, Mois, Jour);
   for i_ind:=0 to TobVente.detail.count-1 do
      begin
      if not VarIsNull(TobVente.Detail[i_ind].GetValue('RUPT')) then
         StRupt := TobVente.Detail[i_ind].GetValue('RUPT')
      else StRupt:='';
      MoisVente:=TobVente.Detail[i_ind].GetValue('MOIS');
      AnneeVente:=TobVente.Detail[i_ind].GetValue('ANNEE');
      // Calcul pour savoir le numéro de la vente, par rapport au mois et à l'année
      // Ceci pour inseré la qté vendue au bonne endroit dans la table GCTMPETATVTE
      noVente:=Mois-MoisVente;
      if AnneeVente<>Annee then noVente:=noVente+nbmois;
      if tri<>'' then
         begin
         if not VarIsNull(TobVente.Detail[i_ind].GetValue('TRI')) then
            StTri :=  TobVente.Detail[i_ind].GetValue('TRI')
         else StTri:='';
         TobT := TobStock.FindFirst(['GZL_RUPT1','GZL_TRI'],[StRupt,StTri],True);
         if TobT<>nil then TobT.PutValue('GZL_QTE'+IntToStr(noVente),TobVente.Detail[i_ind].GetValue('QTEVTE'));
         end
      else
         begin
         TobT := TobStock.FindFirst(['GZL_RUPT1'],[StRupt],True);
         if TobT<>nil then TobT.PutValue('GZL_QTE'+IntToStr(noVente),TobVente.Detail[i_ind].GetValue('QTEVTE'));
         end;
      end;
   TobVente.free;
   end
else Ferme(QLigne);

if TobStock.Detail.Count > 0 then TobStock.InsertDB(nil);
TobStock.free;
end;

procedure TOF_EdEtVte.OnClose ;
begin
Inherited;
ExecuteSQL('DELETE FROM GCTMPETATVTE WHERE GZL_UTILISATEUR = "'+V_PGI.USer+'" ');
end;


Initialization
registerclasses([TOF_EdEtVte]);

end.
