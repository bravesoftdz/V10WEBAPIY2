{***********UNITE*************************************************
Auteur  ...... : Vincent Laroche
Créé le ...... : 10/02/2003
Modifié le ... :   /  /
Description .. : Passage en eAGL
Mots clefs ... :
*****************************************************************}
unit SoldeCpt;

interface

uses SysUtils, WinTypes, WinProcs, Messages, Classes,
{$IFNDEF EAGLSERVER}
Controls, Forms, HStatus,
{$ENDIF}
StdCtrls,
{$IFNDEF EAGLCLIENT}
    {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF}
{$ELSE}
     UTOB,
{$ENDIF}
     Hctrls,
     {$IFDEF VER150}
     variants,
     {$ENDIF}
     ExtCtrls,
     ComCtrls,
     HEnt1,
     uLibExercice,
     utilPGI,       // EstTablePartagee
     utilSais,      // CReinitCumulsMS,
     {$IFDEF MODENT1}
     CPTypeCons,
     CPProcMetier,
     {$ENDIF MODENT1}
     ENT1,uEntCommun;


Type Tab6Double = Array[1..6] Of Double ;
Function MajTotComptes(Setfb : TSetFichierBase ; OdAnal,OkChargemag : Boolean ; CodeExo : String ; Compte : string = ''; Initialise : Boolean=TRUE) : Integer ;
{ ODAnal=TRUE uniquement sur fbJal }
Function MajTotTousComptes(OkChargeMag : Boolean ; CodeExo : String) : Integer ;

implementation

uses
  UProcGen;

var FExoNMoins2 : TExoDate;

function FabriqRqInit (vP,vS,vCodeExo :string ) : string ;
begin
result := '' ;
if (vS<>'P') then result:=vP+'_TOTALDEBIT=0, '+vP+'_TOTALCREDIT=0, ' ;
if vCodeExo='' then
 begin
  result:=result + vP+'_TOTDEBP=0, '+vP+'_TOTCREP=0, '+
  vP+'_TOTDEBE=0, '+vP+'_TOTCREE=0, '+
  vP+'_TOTDEBS=0, '+vP+'_TOTCRES=0 ' ;
  if VP = 'G' then
    Result := Result + ', ' + VP + '_TOTDEBN2 = 0, ' + VP + '_TOTCREN2 = 0 ';
  if (vP<>'J') then
   result:=result +','+vP+'_TOTDEBANO=0, '+vP+'_TOTCREANO=0, '+
   vP+'_TOTDEBANON1=0, '+ vP+'_TOTCREANON1=0 '
 end
 else
  begin
   result:=result+vP+'_TOTDEB'+vS+'=0, '+vP+'_TOTCRE'+vS+'=0 ' ;
   if (vS='E') and (vP<>'J') then
    result:=result+','+vP+'_TOTDEBANO=0, '+vP+'_TOTCREANO=0 ' ;
   if (vS='S') and (vP<>'J') then
    result:=result+','+vP+'_TOTDEBANON1=0, '+vP+'_TOTCREANON1=0 ' ;
  end ;
end;

{=============================================================================}
{***********A.G.L.***********************************************
Auteur  ...... : ?
Créé le ...... : 08/02/2005
Modifié le ... : 10/02/2005
Description .. : - LG - 08/02/2005 - ajout d'un parametre Compte. On ne
Suite ........ : mets au jour que le compte ou l'auxi passer en parametre (
Suite ........ : sert ds la com )
Suite ........ : - gestion l mise a jour des montants en fct de l'exercice que
Suite ........ : l'on mets a jour
Mots clefs ... :
*****************************************************************}
procedure ReinitSolde ( fb,fb1 : TFichierBase ; ODAnal : Boolean ; CodeExo : string ; Compte : string = '') ;
Var StInit : String ;
 lSt : string ;
begin
lSt := '' ;
if CodeExo = '' then lSt:='' else
if CodeExo = GetPrecedent.Code then lSt:='P' else
if CodeExo = GetEnCours.Code then lSt:='E' else
if CodeExo = GetSuivant.Code then lSt:='S' else
if CodeExo = FExoNMoins2.Code then lSt :='N2'; // ajout me pour N-2

  if not EstTablePartagee( fbToTable( fb ) ) then
    begin
    Case Fb Of
      fbGene :begin
               StInit:='UPDATE GENERAUX SET ' ;
               StInit:=StInit + FabriqRqInit('G',lSt,CodeExo) ;
               end ;
      fbAux : begin
               StInit:='UPDATE TIERS SET ' ;
               StInit:=StInit + FabriqRqInit('T',lSt,CodeExo) ;
             end ; // case
      fbAxe1..fbAxe5 : begin
                         StInit:='UPDATE SECTION SET ' ;
                         StInit:=StInit + FabriqRqInit('S',lSt,CodeExo) + ' WHERE S_AXE="'+fbToAxe(fb)+'"' ;
                       end ;
      fbJal  : BEGIN
                StInit:='UPDATE JOURNAL SET ' ;
                StInit:=StInit + FabriqRqInit('J',lSt,CodeExo) ;
               If OdAnal Then
                  BEGIN
                  StInit:=StInit+'WHERE ((J_NATUREJAL="ODA") Or (J_NATUREJAL="ANA")) AND J_AXE="'+fbToAxe(fb1)+'"' ;
                  END
                  Else StInit:=StInit+'WHERE J_NATUREJAL<>"ODA" And J_NATUREJAL<>"ANA"' ;
               END ;
      END ;
    if (Compte<>'') and (fb=fbGene) then StInit:=StInit+' WHERE G_GENERAL="' + Compte + '" ' ;
    if (Compte<>'') and (fb=fbAux) then StInit:=StInit+' WHERE T_AUXILIAIRE="' + Compte + '" ' ;
    if (Compte<>'') and (fb=fbJal) then StInit:=StInit+' AND J_JOURNAL="' + Compte + '" ' ;
    ExecuteSQL(StInit) ;
    end
  else
    // Gestion du multiSociete (restrictions gérées dans la fonction)
    Case Fb Of
      fbGene, fbAux  : CReinitCumulsMS( fb,     Compte, '',           CodeExo, ODAnal , Compte = '' ) ;
      fbAxe1..fbAxe5 : CReinitCumulsMS( fbSect, Compte, fbToAxe(fb),  CodeExo, ODAnal , Compte = '' ) ;
      fbJal          : if ODAnal
                         then CReinitCumulsMS( fbJal,  Compte, fbToAxe(fb1), CodeExo, ODAnal , Compte = '' )
                         else CReinitCumulsMS( fbJal,  Compte, '',           CodeExo, ODAnal , Compte = '' ) ;
      end ;

end ;

{=============================================================================}
Function FabricReqCpt (fb : TFichierBase ; TD,TC : Tab6Double ; D,C : Double ; pCpte : String17; Initialise : Boolean; CodeExo : string) : String ;
Var SQL : String ;
    DD,CD,DP,CP,DE,CE,DS,CS,DA,CA,DA1,CA1,CPTE : String;
    DNMoins2, CNMoins2 : String;
begin
DD:=VariantToSQL(D);     CD:=VariantToSQL(C);
DP:=VariantToSQL(TD[1]); CP:=VariantToSQL(TC[1]);
DE:=VariantToSQL(TD[2]); CE:=VariantToSQL(TC[2]);
DS:=VariantToSQL(TD[3]); CS:=VariantToSQL(TC[3]);
If fb<>fbJal then begin
  DA:=VariantToSQL(TD[4]); CA:=VariantToSQL(TC[4]);
  DA1:=VariantToSQL(TD[5]);CA1:=VariantToSQL(TC[5]);
end;

// GCO - 14/20/2005
DNMoins2 := VariantToSQL(TD[6]);
CNMoins2 := VariantToSQL(TC[6]);

CPTE:='"'+pCpte+'"';
if Initialise then
begin
      Case fb of
         fbGene : BEGIN
                  SQL:='UPDATE GENERAUX SET G_TOTALDEBIT=G_TOTALDEBIT+'+DD+', G_TOTALCREDIT=G_TOTALCREDIT+'+CD+', ' ;
                  SQL:=SQL+'G_TOTDEBE=G_TOTDEBE+'+DE+', G_TOTCREE=G_TOTCREE+'+CE+', G_TOTDEBS=G_TOTDEBS+'+DS+', G_TOTCRES=G_TOTCRES+'+CS+', ' ;
                  SQL:=SQL+'G_TOTDEBP=G_TOTDEBP+'+DP+', G_TOTCREP=G_TOTCREP+'+CP+', ' ;
                  SQL:=SQL+'G_TOTDEBN2=G_TOTDEBN2+'+DNMoins2+', G_TOTCREN2=G_TOTCREN2+'+CNMoins2+', ';
                  SQL:=SQL+'G_TOTDEBANO=G_TOTDEBANO+'+DA+', G_TOTCREANO=G_TOTCREANO+'+CA+', ' ;
                  SQL:=SQL+'G_TOTDEBANON1=G_TOTDEBANON1+'+DA1+', G_TOTCREANON1=G_TOTCREANON1+'+CA1+' WHERE G_GENERAL='+CPTE;
                  END ;
          fbAux : BEGIN
                  SQL:='UPDATE TIERS SET T_TOTALDEBIT=T_TOTALDEBIT+'+DD+', T_TOTALCREDIT=T_TOTALCREDIT+'+CD+', ' ;
                  SQL:=SQL+'T_TOTDEBE=T_TOTDEBE+'+DE+', T_TOTCREE=T_TOTCREE+'+CE+', T_TOTDEBS=T_TOTDEBS+'+DS+', T_TOTCRES=T_TOTCRES+'+CS+', ' ;
                  SQL:=SQL+'T_TOTDEBP=T_TOTDEBP+'+DP+', T_TOTCREP=T_TOTCREP+'+CP+', '  ;
                  SQL:=SQL+'T_TOTDEBANO=T_TOTDEBANO+'+DA+', T_TOTCREANO=T_TOTCREANO+'+CA+', ' ;
                  SQL:=SQL+'T_TOTDEBANON1=T_TOTDEBANON1+'+DA1+', T_TOTCREANON1=T_TOTCREANON1+'+CA1+' WHERE T_AUXILIAIRE='+CPTE;
                  END ;
         fbAxe1..fbAXe5 :
                  BEGIN
                  SQL:='UPDATE SECTION SET S_TOTALDEBIT=S_TOTALDEBIT+'+DD+', S_TOTALCREDIT=S_TOTALCREDIT+'+CD+', ' ;
                  SQL:=SQL+'S_TOTDEBE=S_TOTDEBE+'+DE+', S_TOTCREE=S_TOTCREE+'+CE+', S_TOTDEBS=S_TOTDEBS+'+DS+', S_TOTCRES=S_TOTCRES+'+CS+', ' ;
                  SQL:=SQL+'S_TOTDEBP=S_TOTDEBP+'+DP+', S_TOTCREP=S_TOTCREP+'+CP+', ' ;
                  SQL:=SQL+'S_TOTDEBANO=S_TOTDEBANO+'+DA+', S_TOTCREANO=S_TOTCREANO+'+CA+', ' ;
                  SQL:=SQL+'S_TOTDEBANON1=S_TOTDEBANON1+'+DA1+', S_TOTCREANON1=S_TOTCREANON1+'+CA1+' WHERE S_AXE="'+fbToAxe(fb)+'" AND S_SECTION='+CPTE;
                  END ;
         fbJal : BEGIN
                  SQL:='UPDATE JOURNAL SET J_TOTALDEBIT=J_TOTALDEBIT+'+DD+', J_TOTALCREDIT=J_TOTALCREDIT+'+CD+', ' ;
                  SQL:=SQL+'J_TOTDEBE=J_TOTDEBE+'+DE+', J_TOTCREE=J_TOTCREE+'+CE+', J_TOTDEBS=J_TOTDEBS+'+DS+', J_TOTCRES=J_TOTCRES+'+CS+', ' ;
                  SQL:=SQL+'J_TOTDEBP=J_TOTDEBP+'+DP+', J_TOTCREP=J_TOTCREP+'+CP+' WHERE J_JOURNAL='+CPTE;
                  END ;
         End ;
end
else
begin
     Case fb of
             fbGene : BEGIN
                      SQL :='UPDATE GENERAUX SET G_TOTALDEBIT='+DD+', G_TOTALCREDIT='+CD;

                      if (CodeExo = GetEnCours.Code) or (CodeExo = '') then
                      begin
                          SQL := SQL + ', G_TOTDEBE='+ DE +', G_TOTCREE='+ CE;
                          SQL:=SQL+', G_TOTDEBANO='+DA+', G_TOTCREANO='+CA;
                      end;
                      if (CodeExo = GetSuivant.Code) or (CodeExo = '') then
                      begin
                          SQL := SQL + ', G_TOTDEBS='+ DS +', G_TOTCRES='+ CS;
                          SQL:=SQL+', G_TOTDEBANON1='+DA1+', G_TOTCREANON1='+CA1;
                      end;

                      if (CodeExo = GetPrecedent.Code) or (CodeExo = '') then
                          SQL:=SQL+', G_TOTDEBP='+DP+', G_TOTCREP='+ CP ;

                      if (CodeExo = FExoNMoins2.Code) or  (CodeExo = '') then
                          SQL:=SQL+', G_TOTDEBN2='+DNMoins2+', G_TOTCREN2='+CNMoins2;

                      SQL := SQL + ' WHERE G_GENERAL='+CPTE;
                      END ;
              fbAux : BEGIN
                      SQL:='UPDATE TIERS SET T_TOTALDEBIT='+DD+', T_TOTALCREDIT='+CD;
                      if (CodeExo = GetEnCours.Code) or (CodeExo = '') then
                      begin
                            SQL:=SQL+', T_TOTDEBE='+DE+', T_TOTCREE='+CE;
                            SQL:=SQL+', T_TOTDEBANO='+DA+', T_TOTCREANO='+CA ;
                      end;
                      if (CodeExo = GetSuivant.Code) or (CodeExo = '') then
                      begin
                            SQL:=SQL+', T_TOTDEBS='+DS+', T_TOTCRES='+CS;
                            SQL:=SQL+', T_TOTDEBANON1='+DA1+', T_TOTCREANON1='+CA1;
                      end;
                      if (CodeExo = GetPrecedent.Code) or (CodeExo = '') then
                        SQL:=SQL+', T_TOTDEBP='+DP+', T_TOTCREP='+CP;

                      SQL:=SQL+ ' WHERE T_AUXILIAIRE='+CPTE;
                      END ;
             fbAxe1..fbAXe5 :
                      BEGIN
                      SQL:='UPDATE SECTION SET S_TOTALDEBIT='+DD+', S_TOTALCREDIT='+CD;
                      if (CodeExo = GetEnCours.Code) or (CodeExo = '') then
                      begin
                          SQL:=SQL+', S_TOTDEBE='+DE+', S_TOTCREE='+CE;
                          SQL:=SQL+', S_TOTDEBANO='+DA+', S_TOTCREANO='+CA;
                      end;
                      if (CodeExo = GetSuivant.Code) or (CodeExo = '') then
                      begin
                          SQL:=SQL+', S_TOTDEBS='+DS+', S_TOTCRES='+CS;
                          SQL:=SQL+', S_TOTDEBANON1='+DA1+', S_TOTCREANON1='+CA1;
                      end;
                      if (CodeExo = GetPrecedent.Code) or (CodeExo = '') then
                          SQL:=SQL+', S_TOTDEBP='+DP+', S_TOTCREP='+CP;

                      SQL:=SQL+' WHERE S_AXE="'+fbToAxe(fb)+'" AND S_SECTION='+CPTE;
                      END ;
             fbJal : BEGIN
                      SQL:='UPDATE JOURNAL SET J_TOTALDEBIT='+DD+', J_TOTALCREDIT='+CD;
                      if (CodeExo = GetEnCours.Code) or (CodeExo = '') then
                      SQL:=SQL+', J_TOTDEBE='+DE+', J_TOTCREE='+CE;
                      if (CodeExo = GetSuivant.Code) or (CodeExo = '') then
                      SQL:=SQL+', J_TOTDEBS='+DS+', J_TOTCRES='+CS;
                       if (CodeExo = GetPrecedent.Code) or (CodeExo = '') then
                      SQL:=SQL+', J_TOTDEBP='+DP+', J_TOTCREP='+CP;
                      SQL:=SQL+ ' WHERE J_JOURNAL='+CPTE;
                      END ;
     End ;
end;
Result:=SQL;
END ;


{=============================================================================}
{***********A.G.L.***********************************************
Auteur  ...... : ?
Créé le ...... : 08/02/2005
Modifié le ... : 20/04/2006
Description .. : - LG - 08/02/2005 - ajout d'un parametre Compte. On ne
Suite ........ : mets au jour que le compte ou l'auxi passer en parametre (
Suite ........ : sert ds la com )
Suite ........ : - ajout de la section ds la rq de selection des somme par
Suite ........ : axes
Suite ........ : - ajout d'une condition pour la recherche sur les auxi : on ne 
Suite ........ : prends en compte que les lignes d'ecritures avec des 
Suite ........ : auxiliaires
Suite ........ : 
Suite ........ : SBO 20/04/2006 : Ajout test sur la date comptable quand
Suite ........ : exercice non renseigné pour ne pas ramené les écritures
Suite ........ : plus anciennes que N-2
Mots clefs ... :
*****************************************************************}
Function FabricReqCalc ( fb,fb1 : TFichierBase ; ODAnal : Boolean ; CodeExo : String ; Compte : string = '') : String ;
Var StCalc : String ;
begin
Case Fb Of
  FbGene : StCalc:='SELECT E_GENERAL, E_EXERCICE, ' ;
  FbAux  : StCalc:='SELECT E_AUXILIAIRE, E_EXERCICE, ';
  fbAxe1..fbAxe5 : StCalc:='SELECT Y_SECTION, Y_EXERCICE, ' ;
  fbJal  : BEGIN
           If ODAnal Then StCalc:='SELECT Y_JOURNAL, Y_EXERCICE, '
                     Else StCalc:='SELECT E_JOURNAL, E_EXERCICE, ' ;
           END ;
  END ;
Case fb Of
  fbGene,fbAux : begin
                  StCalc:=StCalc+'sum(E_DEBIT), sum(E_CREDIT), E_ECRANOUVEAU FROM ECRITURE WHERE E_QUALIFPIECE="N" ' ;
                  if (fb=fbAux) then StCalc:=StCalc+' AND E_AUXILIAIRE<>"" ' ;
                  if (Compte<>'') and (fb=fbGene) then StCalc:=StCalc+' AND E_GENERAL="' + Compte + '" ' ;
                  if (Compte<>'') and (fb=fbAux) then StCalc:=StCalc+' AND E_AUXILIAIRE="' + Compte + '" ' ;
                  if (CodeExo<>'')
                    then StCalc:=StCalc + 'AND E_EXERCICE="' + CodeExo + '" '
                    else StCalc:=StCalc + 'AND E_DATECOMPTABLE >= "' + UsDateTime( FExoNMoins2.Deb ) + '" ' ;
                 end ;
  fbAxe1..fbAxe5 : begin
                  //SG6 07.03.05 Gestion analyutique croisaxe
                  if not GetAnaCroisaxe then
                    StCalc:=StCalc+'sum(Y_DEBIT), sum(Y_CREDIT), Y_ECRANOUVEAU FROM ANALYTIQ WHERE Y_QUALIFPIECE="N" ' +
                    ' AND Y_AXE="'+fbToAxe(fb)+'"'
                  else
                    StCalc:=StCalc+'sum(Y_DEBIT), sum(Y_CREDIT), Y_ECRANOUVEAU, Y_SOUSPLAN1, Y_SOUSPLAN2, Y_SOUSPLAN3, Y_SOUSPLAN4, Y_SOUSPLAN5 FROM ANALYTIQ WHERE Y_QUALIFPIECE="N" ' +
                    ' AND Y_AXE="'+fbToAxe(fb)+'"';

                    if (CodeExo<>'')
                      then StCalc := StCalc + 'AND Y_EXERCICE="' + CodeExo + '" '
                      else StCalc := StCalc + 'AND Y_DATECOMPTABLE >= "' + UsDateTime( FExoNMoins2.Deb ) + '" ' ;
                    if (Compte<>'') then StCalc := StCalc +' AND Y_SECTION="' + Compte + '" ' ;
                   end ;
  fbJal : BEGIN
          If ODAnal Then
             BEGIN
             StCalc:=StCalc+'sum(Y_DEBIT), sum(Y_CREDIT) FROM ANALYTIQ WHERE Y_QUALIFPIECE="N" AND Y_TYPEANALYTIQUE="X" ' ;
             //SG6 07.03.05 Gestion ana croisaxe
             if not GetAnaCroisaxe then StCalc:=StCalc+'AND Y_AXE="'+fbToAxe(fb1)+'"' ;

             if (CodeExo<>'')
               then StCalc:=StCalc + 'AND Y_EXERCICE="' + CodeExo + '" '
               else StCalc := StCalc + 'AND Y_DATECOMPTABLE >= "' + UsDateTime( FExoNMoins2.Deb ) + '" ' ;
             if (Compte<>'') and (fb=fbJal) then StCalc:=StCalc+' AND Y_JOURNAL="' + Compte + '" ' ;
             END Else
             BEGIN
             StCalc:=StCalc+'sum(E_DEBIT), sum(E_CREDIT) FROM ECRITURE WHERE E_QUALIFPIECE="N" ' ;
             if (CodeExo<>'')
               then StCalc:=StCalc + 'AND E_EXERCICE="' + CodeExo + '" '
               else StCalc:=StCalc + 'AND E_DATECOMPTABLE >= "' + UsDateTime( FExoNMoins2.Deb ) + '" ' ;
             if (Compte<>'') and (fb=fbJal) then StCalc:=StCalc+' AND E_JOURNAL="' + Compte + '" ' ;
             END ;
          END ;
  END ;
Case Fb Of
  FbGene : StCalc:=StCalc + ' GROUP BY E_GENERAL,E_EXERCICE,E_ECRANOUVEAU ' ;
  FbAux  : StCalc:=StCalc + ' GROUP BY E_AUXILIAIRE,E_EXERCICE,E_ECRANOUVEAU ' ;
  fbAxe1..fbAxe5 : begin
                   //SG6 07.03.05 Gestion analyutique croisaxe
                   if not GetAnaCroisaxe then
                     StCalc:=StCalc+ ' GROUP BY Y_SECTION,Y_EXERCICE,Y_ECRANOUVEAU'
                   else
                     StCalc:=StCalc+ ' GROUP BY Y_SECTION,Y_EXERCICE,Y_ECRANOUVEAU, Y_SOUSPLAN1, Y_SOUSPLAN2, Y_SOUSPLAN3, Y_SOUSPLAN4, Y_SOUSPLAN5'

                   end;
  FbJal  : BEGIN
           If ODAnal Then StCalc:=StCalc+'GROUP BY Y_JOURNAL,Y_EXERCICE '
                     Else StCalc:=StCalc+'GROUP BY E_JOURNAL,E_EXERCICE ' ;
           END ;
  END ;
Result:=StCalc ;
end ;

{=============================================================================}
{***********A.G.L.***********************************************
Auteur  ...... : ?
Créé le ...... : 08/02/2005
Modifié le ... :   /  /
Description .. : - LG - 08/02/2005 - ajout d'un parametre Compte. On ne
Suite ........ : mets au jour que le compte ou l'auxi passer en parametre (
Suite ........ : sert ds la com )
Mots clefs ... : si Initialise=FALSE cas Comsx pour éviter update=0  des totaux
               : la première fois total = montant, deuxième fois total= total+montant
               : si Initialise=TRUE, update à 0 des totaux après total = total+montant
*****************************************************************}
Function MajTotCompte(fb,fb1 : TFichierBase ; OdAnal : Boolean ; CodeExo : String ; Compte : string = ''; Initialise : Boolean=TRUE) : Integer ;
Var StCalc : String ;
    QCalc : TQuery ;
    SQLCpt : String;
    Cpt : String17 ;
    TD,TC : Double ;
    TotD,TotC : Tab6Double ;
    TTD,TTC : Tab6Double ;
    Exo : String3 ;
    AN : String3 ;
    i,j : Integer ;
    chp : variant;
    sousPlan : String ;
    PlusUneFois : Boolean;
  Label 1 ;
begin
  Result:=0 ;

  if Initialise then ReinitSolde(Fb,Fb1 ,ODAnal,CodeExo,Compte) ;
  Fillchar(TotD,SizeOf(TotD),#0) ; Fillchar(TotC,SizeOf(TotC),#0) ;
  StCalc:=FabricReqCalc(Fb,Fb1,ODAnal,CodeExo,Compte) ; QCalc:=OpenSQL(StCalc,TRUE,-1,'',true) ;

  if (QCalc.EOF) and (not Initialise) then // ajout me pour initialiser les comptes qui n'ont pas de mvts
      ReinitSolde(Fb,Fb1 ,ODAnal,CodeExo,Compte) ;

  {$IFNDEF EAGLSERVER}
  if Compte = '' then InitMove(RecordsCount(QCalc),'Calculs en cours...') ;
  {$ENDIF}
  PlusUneFois := FALSE;
  Try
    While Not QCalc.EOF Do
    BEGIN
      AN:='N' ;
      {$IFNDEF EAGLSERVER}
      if Compte = '' then MoveCur(FALSE) ;
      {$ENDIF}
      Cpt:=QCalc.Fields[0].AsString ;
      Exo:=QCalc.Fields[1].AsString ;
      TD:=Arrondi(QCalc.Fields[2].AsFloat,V_PGI.OkDecV) ;
      TC:=Arrondi(QCalc.Fields[3].AsFloat,V_PGI.OkDecV) ;
      If fb<>fbJal Then AN:=QCalc.Fields[4].AsString ;

      If (TD<>0) Or (TC<>0) Then
      BEGIN
         i:=0 ;
         Fillchar(TTD,SizeOf(TTD),#0) ; Fillchar(TTC,SizeOf(TTC),#0) ;
         If Exo = GetPrecedent.Code Then i:=1 Else
            If Exo = GetEnCours.Code Then i:=2 Else
               If Exo = GetSuivant.Code Then i:=3 else
                 if Exo = FExoNMoins2.Code then i:=6; // si i=0 Alors cumul est sur N-2 au maximum

         If i=0 Then
           BEGIN
           TotD[4]:=Arrondi(TotD[4]+TD,V_PGI.OkDecV) ;
           TotC[4]:=Arrondi(TotC[4]+TD,V_PGI.OkDecV) ;
           END Else
           BEGIN
           TotD[i]:=Arrondi(TotD[i]+TD,V_PGI.OkDecV) ;
           TotC[i]:=Arrondi(TotC[i]+TD,V_PGI.OkDecV) ;
           END ;
         If (TD<>0) And (i<>0) Then TTD[i]:=TD ; If (TC<>0) And (i<>0) Then TTC[i]:=TC ;
         If (fb<>fbJal) Then
            BEGIN
            If (AN<>'N') And (i=2) Then
               BEGIN
               If TD<>0 Then TTD[4]:=TD ; If TC<>0 Then TTC[4]:=TC ;
               END ;
            If (AN<>'N') And (i=3) Then
               BEGIN
               If TD<>0 Then TTD[5]:=TD ; If TC<>0 Then TTC[5]:=TC ;
               END ;
            { Pas de recalcul pour X_TOTALDEBIT et X_TOTALCREDIT si Cumul d'a-nouveau sur un autre exercice que l'en cours }
            If (i<>2) Then If AN<>'N' Then BEGIN TD:=0 ; TC:=0 ; END ;
            END ;
    //     if V_PGI.ExoV8.Code<>'' then Stcalc:=Stcalc+' And E_DATECOMPTABLE>="'+UsDateTime(V_PGI.ExoV8.Deb)+'" ' ;
         If (i<=1) or (i=6) Then BEGIN TD:=0 ; TC:=0 ; END ;

         {JP 30/11/07 : Mise à jour proprement dite des cumuls}
         BeginTrans ;
         try
           //SG6 07.03.05 Gestion analytique croisaxe
           if (not GetAnaCroisaxe) or (Fb in [fbJal,fbGene,fbAux]) then
           begin

             if not EstTablePartagee( fbToTable( Fb ) ) then
               begin
                   // ajout me dans le cas inilialise à FALSE, pour ne pas écraser avec group by E_ECRANOUVEAU les totaux des comptes
               if (not Initialise) and PlusUneFois then
                   SQLCpt := FabricReqCpt(Fb,TTD,TTC,TD,TC,Cpt, True, Exo)
               else
                   SQLCpt := FabricReqCpt(Fb,TTD,TTC,TD,TC,Cpt, Initialise, Exo);
               ExecuteSQL(SQLCpt);
               end
             else
               // Gestion du multiSociete (restrictions gérées dans la fonction)
               Case Fb Of
                 fbGene, fbAux, fbJal  :
                 begin
                   if (not Initialise) and PlusUneFois then
                     CMajTotCumulsMS( fb, Cpt, '', TD, TC, TTD, TTC, FALSE )
                   else
                     CMajTotCumulsMS( fb, Cpt, '', TD, TC, TTD, TTC, not initialise ) ;
                 end;
                 fbAxe1..fbAxe5,fbSect :
                 begin
                   if (not Initialise) and PlusUneFois then
                     CMajTotCumulsMS( fbSect, Cpt, fbToAxe( fb ), TD, TC, TTD, TTC, FALSE )
                   else
                     CMajTotCumulsMS( fbSect, Cpt, fbToAxe( fb ), TD, TC, TTD, TTC, not initialise ) ;
                 end;
               end ;

           end
           else
           begin
             for j := 1 to 5 do
             begin
               chp := QCalc.FindField('Y_SOUSPLAN' + IntToStr(j)).AsVariant;
               if VarIsNull(chp) then continue;

               sousPlan := QCalc.FindField('Y_SOUSPLAN' + IntToStr(j)).AsString ;
               if sousPlan = '' then continue;

               if not EstTablePartagee( 'SECTION' ) then
                 begin
                 SQLCpt := FabricReqCpt(AxeToFb('A' + IntToStr(j)),TTD,TTC,TD,TC,sousPlan,Initialise, Exo);
                 ExecuteSQL(SQLCpt);
                 end
               else
                 // Gestion du multiSociete (restrictions gérées dans la fonction, cf utilsais.pas)
                 CMajTotCumulsMS( fbSect, sousPlan, 'A' + IntToStr(j), TD, TC, TTD, TTC, not Initialise ) ;

             end;
           end;
           CommitTrans;
         except
           on E : Exception do RollBackDiscret;
         end;

       END ;
      QCalc.Next ;
      PlusUneFois := TRUE;
    END ;
  finally
    Ferme(QCalc) ;
  end;
  {$IFNDEF EAGLSERVER}
  if Compte = '' then FiniMove ;
  {$ENDIF}

  i:=0 ;
  If Arrondi(TotD[1]-TotC[1],V_PGI.OkDecV)<>0 Then Result:=i ;
  If Arrondi(TotD[2]-TotC[2],V_PGI.OkDecV)<>0 Then Result:=Result+(10*i) ;
  If Arrondi(TotD[3]-TotC[3],V_PGI.OkDecV)<>0 Then Result:=Result+(100*i) ;
  If Arrondi(TotD[4]-TotC[4],V_PGI.OkDecV)<>0 Then Result:=Result+(1000*i) ;
end ;

{=============================================================================}
{***********A.G.L.***********************************************
Auteur  ...... : ?
Créé le ...... : 08/02/2005
Modifié le ... :   /  /
Description .. : - LG - 08/02/2005 - ajout d'un parametre Compte. On ne
Suite ........ : mets au jour que le compte ou l'auxi passer en parametre (
Suite ........ : sert ds la com )
Mots clefs ... :
*****************************************************************}
Function MajTotComptes(Setfb : TSetFichierBase ; OdAnal,OkChargemag : Boolean ; CodeExo : String ; Compte : string = ''; Initialise : Boolean=TRUE) : Integer ;
begin
Result:=0 ;

{$IFNDEF NOVH}
FExoNMoins2 := CGetExerciceNMoins2;
if ((CodeExo<>'') and (CodeExo<>VH^.Precedent.Code) and (CodeExo<>VH^.EnCours.Code) and (CodeExo<>VH^.Suivant.Code) and
    (CodeExo <> FExoNMoins2.Code))   // ajout me pour N-2
   or ((CodeExo<>'') and (CodeExo = FExoNMoins2.Code) and ((fbGene in Setfb)=FALSE)) then exit ;
If Okchargemag Then ChargeMagHalley ;
{$ELSE}
FExoNMoins2 := GetNMoins2 ;
if ((CodeExo<>'') and (CodeExo<> GetPrecedent.Code) and (CodeExo<> GetEnCours.Code) and (CodeExo<>GetSuivant.Code) and
    (CodeExo <> FExoNMoins2.Code))   // ajout me pour N-2
   or ((CodeExo<>'') and (CodeExo = FExoNMoins2.Code) and ((fbGene in Setfb)=FALSE)) then exit ;
{$ENDIF}

If fbGene in Setfb Then BEGIN Result:=MajTotCompte(fbGene,fbgene,FALSE,CodeExo,Compte, Initialise) ; (*RecalculTotPointeNew('') ;*) END ;
If fbAux in Setfb Then MajTotCompte(fbAux,fbAux,FALSE,CodeExo,Compte,Initialise) ;
If (fbAxe1 in Setfb) Or (fbSect in Setfb) Then MajTotCompte(fbAxe1,fbAxe1,FALSE,CodeExo,Compte, Initialise) ;
If (fbAxe2 in Setfb) Or (fbSect in Setfb) Then MajTotCompte(fbAxe2,fbAxe2,FALSE,CodeExo,Compte, Initialise) ;
If (fbAxe3 in Setfb) Or (fbSect in Setfb) Then MajTotCompte(fbAxe3,fbAxe3,FALSE,CodeExo,Compte, Initialise) ;
If (fbAxe4 in Setfb) Or (fbSect in Setfb) Then MajTotCompte(fbAxe4,fbAxe4,FALSE,CodeExo,Compte, Initialise) ;
If (fbAxe5 in Setfb) Or (fbSect in Setfb) Then MajTotCompte(fbAxe5,fbAxe5,FALSE,CodeExo,Compte, Initialise) ;
If fbJal in Setfb Then
   BEGIN
   If ODAnal Then
      BEGIN
      MajTotCompte(fbjal,fbAxe1,TRUE,CodeExo,Compte, Initialise) ;
      MajTotCompte(fbjal,fbAxe2,TRUE,CodeExo,Compte, Initialise) ;
      MajTotCompte(fbjal,fbAxe3,TRUE,CodeExo,Compte, Initialise) ;
      MajTotCompte(fbjal,fbAxe4,TRUE,CodeExo,Compte, Initialise) ;
      MajTotCompte(fbjal,fbAxe5,TRUE,CodeExo,Compte,Initialise) ;
      END Else MajTotCompte(fbJal,fbJal,OdAnal,CodeExo,Compte, Initialise) ;
   END ;
end ;

{=============================================================================}
Function MajTotTousComptes(OkChargeMag : Boolean ; CodeExo : String) : Integer ;
Var Sfb : TSetFichierBase ;
begin
//Result:=0 ;
{$IFNDEF NOVH}
If OkChargeMag Then ChargeMagHalley ;
{$ENDIF}
Sfb:=[fbGene,fbAux,fbsect,fbJal] ;
Result:=MajTotComptes(Sfb,FALSE,FALSE,CodeExo) ;
Sfb:=[fbJal] ; MajTotComptes(SFb,TRUE,FALSE,CodeExo) ;
end ;


end.
