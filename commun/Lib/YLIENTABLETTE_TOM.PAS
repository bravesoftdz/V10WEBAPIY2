{***********UNITE*************************************************
Auteur  ...... :                                                                  
Créé le ...... : 10/05/2005
Modifié le ... :   /  /
Description .. : Source TOM de la TABLE : YLIENTABLETTE (YLIENTABLETTE)
Mots clefs ... : TOM;YLIENTABLETTE
*****************************************************************}
Unit YLIENTABLETTE_TOM ;

Interface
                                                          

Uses StdCtrls, 
     Controls, 
     Classes, 
{$IFNDEF EAGLCLIENT}
     FE_Main,
     db,
     {$IFNDEF DBXPRESS}dbtables{BDE},{$ELSE}uDbxDataSet,{$ENDIF} 
     Fiche, 
     FichList, 
{$else}
    MaineAGL,
     eFiche,
     eFichList, 
{$ENDIF}
     forms, 
     sysutils, 
     ComCtrls,
     HCtrls, 
     HEnt1, 
     HMsgBox, 
     UTOM, 
     UTob,
     ParamSoc,
     EntGC
     ;

Type
  TOM_YLIENTABLETTE = Class (TOM)
    procedure OnNewRecord                ; override ;
    procedure OnDeleteRecord             ; override ;
    procedure OnUpdateRecord             ; override ;
    procedure OnAfterUpdateRecord        ; override ;
    procedure OnAfterDeleteRecord        ; override ;
    procedure OnLoadRecord               ; override ;
    procedure OnChangeField ( F: TField) ; override ;
    procedure OnArgument ( S: String )   ; override ;
    procedure OnClose                    ; override ;
    procedure OnCancelRecord             ; override ;
    end ;
// mis AGL Procedure MajTablette(NomTablette:string);
Procedure YYLanceFiche_YYLienTablette(Range,Argument:string);
const
  TexteMessage: array[1..6] of string  = (
          {1}         'Le code origine doit être différent du code destination.'
          {2}         ,'Vous devez renseigner les 2 codes.'
          {3}         ,'Ce couple existe déjà.'
          {4}         ,'Le code destinataire (ou le code origine) existe déjà dans le paramétrage.Correspondance impossible.'
          {5}         ,'Le code de vos tablettes ne sont pas sur la même longueur.Il faudra saisir la tablette origine sur 3c seulement.'
          {6}         ,'Vos modifications seront actives à la prochaine connexion.#13#10 Si vous avez créé de nouvelles correspondances, vous devez#13#10 modifier la tablette origine afin que les valeurs soient alignées#13#10 dans la tablette destination.'
                      );

Implementation

procedure TOM_YLIENTABLETTE.OnNewRecord ;
begin
  Inherited ;
end ;

procedure TOM_YLIENTABLETTE.OnDeleteRecord ;
begin
  Inherited ;
end ;

procedure TOM_YLIENTABLETTE.OnUpdateRecord ;
var stdest, stori : string;
begin
  Inherited ;
  if (GetCOntrolText('YLT_ORIGINETAB')='')
    or (GetCOntrolText('YLT_DESTINATIONTAB')='') then
    begin  //code indetique en origine et destination.. onr efsue
    LastError := 2;
    LastErrorMsg :=TexteMessage[2];
    exit;
    end;
  if (GetCOntrolText('YLT_ORIGINETAB')<>'')
    and (GetCOntrolText('YLT_DESTINATIONTAB')=GetCOntrolText('YLT_ORIGINETAB')) then
    begin  //code indetique en origine et destination.. onr efsue
    LastError := 1;
    LastErrorMsg :=TexteMessage[1] ;
    exit;
    end;
  If ExisteSql ('SELECT YLT_ORIGINETAB FROM YLIENTABLETTE WHERE YLT_ORIGINETAB="'+
      GetCOntrolText('YLT_DESTINATIONTAB') +
      '" and YLT_DESTINATIONTAB="' +
      GetCOntrolText('YLT_ORIGINETAB') + '"')
      then begin
      LastError := 3;
      LastErrorMsg :=TexteMessage[3] ;
      exit;
      end;
  If ExisteSql ('SELECT YLT_ORIGINETAB FROM YLIENTABLETTE WHERE((YLT_ORIGINETAB="'+
      GetCOntrolText('YLT_DESTINATIONTAB') +
      '" or YLT_DESTINATIONTAB="' +
      GetCOntrolText('YLT_DESTINATIONTAB') + '") AND YLT_ORIGINETAB<>"' +
      GetCOntrolText('YLT_ORIGINETAB') + '")' +
      'OR (YLT_DESTINATIONTAB="' + GetCOntrolText('YLT_ORIGINETAB') + '")')
      then begin   // le code destiantion ne doit pas déjà exister dans la table
            //le code origine ne doit pas existait en destiantion dans la table
            //sinon pb en synchro
      LastError := 4;
      LastErrorMsg :=TexteMessage[4] ;
      exit;
      end;
  stDest := GetCOntrolText('YLT_DESTINATIONTAB');
  stOri :=GetCOntrolText('YLT_ORIGINETAB');
  if V_PGI.DECombos[TTToNum(stdest)].Prefixe   <>  V_PGI.DECombos[TTToNum(stOri)].Prefixe 
      then begin   //si copie choixcode (CC) dans choixext(YX) risque pb, car 3c dans 6c
      PgiInfo(TexteMessage[5]) ;
      end;

end ;

procedure TOM_YLIENTABLETTE.OnAfterUpdateRecord ;
begin
  Inherited ;
end ;

procedure TOM_YLIENTABLETTE.OnAfterDeleteRecord ;
begin
  Inherited ;
end ;

procedure TOM_YLIENTABLETTE.OnLoadRecord ;
begin
  Inherited ;
end ;

procedure TOM_YLIENTABLETTE.OnChangeField ( F: TField ) ;
begin
  Inherited ;
end ;

procedure TOM_YLIENTABLETTE.OnArgument ( S: String ) ;
Var stArg : string;
begin
  Inherited ;
  SetControlVisible ('BRECHERCHER',False);
  SetControlVisible ('BPARAMLISTE',False);
  SetControlVisible ('BDUPLIQUER',False);
  stArg :='';
    // on fabrique la condition plus d'affichage des tablettes possibles en correspondance
    // pour mémoire @A=affaire, @RES= table ressource, @G = Gescom
    // @GRC = GRC, @C= compta, @P=paie, @S=stock
    //.....
  If CtxGescom In V_PGI.PGIContexte then
     begin
     stArg := 'AND (CO_ABREGE like "%@G%"';
     if GetParamSocSecur('SO_RTGESTINFOS004',false) then stArg := stArg +' OR CO_ABREGE like "%@004%"';
     if GetParamSocSecur('SO_RTGESTINFOS008',false) then stArg := stArg +' OR CO_ABREGE like "%@008%"';
     end;
  If CtxGcAff In V_PGI.PGIContexte then stArg := stArg +' OR CO_ABREGE like "%@A%" OR CO_ABREGE like "%@RES%"';
  If CtxAffaire in V_Pgi.PgiContexte then stArg := stArg +' OR CO_ABREGE like "%@A%" OR CO_ABREGE like "%@RES%"';
  if (GetParamsocSecur('SO_AFLIENDP',false) =true)  then  stArg := stArg +' OR CO_ABREGE like "%@D%"';//mcd 17/10/2005 pour table DP
{$IFDEF GIGI} //mcd 02/11/2006 ?? perte de ces lignes????
  if GetParamsocSecur ('SO_AFRTSUSPECT',false) then starg:= starg + ' OR CO_ABREGE like "%@SUS"';//mcd 19/06/06 pour tbaletteprospect
    //mcd si les tablettes libres suspects veulent être utilisés dans d'autres produits
    //il faut faire la ligne ci-dessus en ¤else, sans test du paramsoc qui n'est visible qu'en GIGI
    // et aller dasn UtofSuspect_mul, pour ne plus mettre en ifdef GIGI le changement de la tof ancetre
{$endif}
{$ifdef STK}
  stArg := stArg +' OR CO_ABREGE like "%@S%"';
{$endif}
  If CtxPaie In V_PGI.PGIContexte then  stArg := 'AND (CO_ABREGE like "%@P%" OR CO_ABREGE like "%@RES%"';
    // ajouter ici les autres contextes..
  if GetParamSocSecur('SO_RTGESTINFOS003',false) then stArg := stArg +' OR CO_ABREGE like "%@003%"';
  if VH_GC.SAVSeria and (GetParamSoc('SO_RTGESTINFOS005')) then stArg := stArg +' OR CO_ABREGE like "%@005%"';
{$IFDEF GPAO}
  if GetParamSocSecur('SO_RTGESTINFOS007',False) then stArg := stArg +' OR CO_ABREGE like "%@007%"';
{$ENDIF !GPAO}
    
    // ajouter ici en fct du contexte, la condition plus de la table
    //le contexte GRC doit être en dernier car ferme la ( de la requête
  if ctxGRC in V_PGI.PGIContexte then
     begin
     if GetParamSocSecur('SO_RTGESTINFOS001',false) then stArg := stArg +' OR CO_ABREGE like "%@001%"';
     if GetParamSocSecur('SO_RTGESTINFOS002',false) then stArg := stArg +' OR CO_ABREGE like "%@002%"';
     if GetParamSocSecur('SO_RTGESTINFOS006',false) then stArg := stArg +' OR CO_ABREGE like "%@006%"';
     starg := stArg +  ' OR CO_ABREGE like "%@GRC%")'
     end
   else stArg :=stArg + ')';
  SetControlProperty ('YLT_ORIGINETAB', 'Plus', StArg);
  SetControlProperty ('YLT_DESTINATIONTAB', 'Plus', StArg);
end ;

procedure TOM_YLIENTABLETTE.OnClose ;
begin
  Inherited ;
  PGIBox(TraduireMemoire(TexteMessage[6]), ecran.caption);
end ;

procedure TOM_YLIENTABLETTE.OnCancelRecord ;
begin
  Inherited ;
end ;

Procedure YYLanceFiche_YYLienTablette(Range,Argument:string);
begin
AGLLanceFiche ('YY','YYLIENTABLETTE',Range,'',Argument);
end;

(* Mmis dans agl 148
  //fct qui sera appelé ou devra être dans m'agl
  //fait la synchro de la table passée avec les tables destinatrice
  //si existe dans le paramétrage des synchro à faire
  //n'appeler cette ftc, que si des modif ont été faite
Procedure MajTablette(NomTablette:string);
Var TobDet, TobTmp, TobDiff, TobDiffDet: tob;
 QQ :tquery;
 ii,jj : integer;
 St ,TabOri,TabDes, PrefOri,PrefDes, NomTable, NomTableDest,Sql: string;
begin
  NomtableDest :='';
  Nomtable :='';
  if ExisteSql ('SELECT YLT_ORIGINETAB FROM YLIENTABLETTE WHERE YLT_DESTINATIONTAB="'
    + NomTablette + '"') then
    begin
    PgiInfo ('Cette tablette est utilisée en destination dans votre paramétrage. Ne va donc plus correspondre à la tablette origine');
    Exit;
    end;
  TobTmp := Tob.create ('tab origine',Nil, -1);
  TobDet := Tob.create ('les liens',Nil,-1);
  Sql := 'SELECT YLT_DESTINATIONTAB FROM YLIENTABLETTE WHERE YLT_ORIGINETAB="'+ NomTablette + '"';
  Tobdet.loaddetailFromSql (Sql);
  if (Tobdet.detail.count >0)
    then begin
         //on récupère type de table et préfixe pour tablette origine
    St := NomTablette;
    jj := TTToNum(St);
    if jj > 0 then
      begin
      TabOri :=  V_PGI.DECombos[jj].Prefixe;
      PrefOri :=  V_PGI.DECombos[jj].tipe;
      end;
       //on charge la tablette qui vient de changer
    If TabOri ='CC' then NomTable:='CHOIXCOD'   //co : commun, pas lieu d'être dans les paramétrages
     else NomTable:='CHOIXEXT';
        //il faut tout prendre car synchro
    Sql := 'SELECT * FROM ' +NomTable + ' WHERE '+ tabOri+'_TYPE="'+PrefOri+ '"';
    QQ := OpenSQL (sql, TRUE);
    TobTmp.loaddetailDB (NomTable,'','',QQ,false,false);
    Ferme(QQ);
    end;
 if (TobTmp.detail.count >0) then
 begin   //il y a des info dans la tablette origine  ..
         //on copie dans toutes les tablettes destinations
   For ii:=0 to Tobdet.detail.count-1 do
      begin
          //on récupère type de table et préfixe pour tablette  destination
        St := TobDet.detail[ii].getvalue ('YLT_DESTINATIONTAB');
        jj := TTToNum(St);
        if jj > 0 then
          begin
          TabDes :=  V_PGI.DECombos[jj].Prefixe;
          PrefDes :=  V_PGI.DECombos[jj].tipe;
          end;
        If TabDes ='CC' then NomTableDest:='CHOIXCOD'   //co : commun, pas lieu d'être dans les paramétrages
         else NomTableDest:='CHOIXEXT';
            //on détruit les velurs de la tablette destination
        ExecuteSql ('DELETE FROM ' +NomTableDest + ' WHERE '+
            tabDes+'_TYPE="'+PrefDes+ '"' );
        If Nomtable=NomTableDest then
          begin //même table, il suffit de changer le type
          For jj:=0 to TobTmp.detail.count-1 do
            begin
            TobTmp.detail[jj].putValue(TabDes+'_TYPE',Prefdes);
            end;
          TobTmp.InsertorUpdateDb(False);
//   ??? faut il faire un   AvertirTable(NomTableDEst) ?? pour Ok en mémoire ???
          end
        else begin //pas même table, il faut faire les requête sur la nouvelle table
          TobDiff := Tob.Create ('',Nil,-1);
          For jj:=0 to TobTmp.detail.count-1 do
            begin
            TobDiffDet := tob.Create (NomTableDest ,TobDIff,-1);
            TobDiffDet.putValue(TabDes+'_TYPE',Prefdes);
            TobDiffDet.putValue(TabDes+'_LIBRE',TobTmp.detail[jj].getvalue(TabOri+'_LIBRE'));
            TobDiffDet.putValue(TabDes+'_CODE',TobTmp.detail[jj].getvalue(TabOri+'_CODE'));
            TobDiffDet.putValue(TabDes+'_ABREGE',TobTmp.detail[jj].getvalue(TabOri+'_ABREGE'));
            TobDiffDet.putValue(TabDes+'_LIBELLE',TobTmp.detail[jj].getvalue(TabOri+'_LIBELLE'));
            end;
          TobDiff.InsertorUpdateDb(False);
//   ??? faut il faire un   AvertirTable(NomTableDEst) ?? pour Ok en mémoire ???
          TobDiff.free;
          end;
      end;
    end;
TobTmp.free;
TobDet.free;
end;  *)

Initialization
  registerclasses ( [ TOM_YLIENTABLETTE ] ) ; 
end.
