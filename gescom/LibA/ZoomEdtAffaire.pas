unit ZoomEdtAffaire;

interface


uses Hent1,sysutils, forms, controls,Hctrls, FactComm, Facture,
{$IFDEF EAGLCLIENT}
     MaineAGL,
{$ELSE}
     {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}
{$ENDIF}
{$IFDEF BTP}
     BTPUtil, BTFactImprTOB,
{$ENDIF}
{$IFDEF GRC}
     ZoomEdtProspect,
{$ENDIF}
     TiersUtil,ActiviteUtil,AFActivite,Entgc,UTob,Classes,uEntCommun;

Procedure AFZoomStringToCleDoc ( StA : String ; Var CleDoc : R_CleDoc ) ;
Procedure AFZoomEdt(Qui : Integer ; Quoi : String) ;
Procedure AFZoomEdtEtat(Quoi : String) ;
function UsDateToStd( StUS : string ) : string ;
//var TheImpressionViaTOB : TImprPieceViaTOB;

implementation

function UsDateToStd( StUS : string ) : string ;
var OldSDF : string ;
    DD : TDateTime ;
begin
OldSDF :=   ShortDateFormat ;
ShortDateFormat:='yyyyMMdd' ;
DD:=StrToDate (StUS) ;
ShortDateFormat:=  OldSDF;
result:=DateToSTr(DD);
end;

Function TypeTiers (Code: String ): String;
Var Q : TQuery;
    NomChamp,ChampRetourne : String;
Begin
Result:='';
NomChamp := 'T_AUXILIAIRE'; ChampRetourne :='T_NATUREAUXI';
Q:=OPENSQL('SELECT '+ ChampRetourne +' From TIERS WHERE '+ NomChamp +'="'+Code+'"',True);
If Not Q.EOF Then result:='T_NATUREAUXI='+Q.Fields[0].AsString;
Ferme(Q);
End;

Procedure AFZoomEdtEtat(Quoi : String) ;
var i,Qui: Integer ;
begin
  i := Pos(';',Quoi);
  Qui := StrToInt (Copy (Quoi,1,i-1));
  Quoi:=Copy(Quoi,i+1,Length(Quoi)-i) ;
{$IFDEF GRC}
  case Qui Of    { valeurs définies dans tablette TTZOOMEDT }
    500..599 : RTZoomEdt(Qui,Quoi) ;
  end;
{$ENDIF}  
  AFZoomEdt(Qui,Quoi) ;
end;


Procedure AFZoomEdt(Qui : Integer ; Quoi : String) ;
Var St,St2,St3,St4 ,Str, Znum: String ;
    CleDoc : R_CleDoc;
    dDate : TDateTime;
BEGIN

Case Qui Of
  300:  BEGIN      // article
        V_PGI.DispatchTT (7,taConsult ,Quoi, '','');
        END;
  301:  BEGIN      // tiers
{$IFDEF BTP}
        V_PGI.DispatchTT (8,taConsult ,Quoi, '',TypeTiers(Quoi));
{$ELSE}
        V_PGI.DispatchTT (8,taConsult ,TiersAuxiliaire(Quoi,True), '','');
{$ENDIF}
        END;
  302:  BEGIN      // Pièces
        FillChar(CleDoc,Sizeof(CleDoc),#0) ;
        AFZoomStringToCleDoc(Quoi,CleDoc) ;
        SaisiePiece(CleDoc,taConsult) ;
        END;
  700:  BEGIN      // affaire
        V_PGI.DispatchTT (5,taConsult ,Quoi, '','');
        END;
  701:  BEGIN      // ressource
        V_PGI.DispatchTT (6,taConsult ,Quoi, '','');
        END;
  702:  BEGIN      // activité par ressource
        St := ReadTokenSt(Quoi); // Ressource
        St2:= ReadTokenSt(Quoi);  // Date
        if (St = '') or (St2 = '') then exit;
        dDate:=StrToDate(USDateToStd(st2));
        if ((dDate=0) or (dDate=iDate1900) or (dDate=iDate2099)) then exit;
        AFCreerActiviteModale(tsaRess, tacGlobal, 'REA', St, '', '', dDate  );
        END;
  703:  BEGIN      // activité  par affaire (sur atb)
        str := ReadTokenPipe(Quoi,';/;') ;
          // double clic avec shift ==> acces factures
        if (ssShift in V_PGI.FZoomKeyClick) and (str<>'') then
         begin
         St := ReadTokenSt(Quoi);  // Nature peice
         St2 := ReadTokenSt(Quoi); // Souche
         St3 := ReadTokenSt(Quoi); // n°
         St4 := ReadTokenSt(Quoi); // date
         if (St = '') or (St2 = '') or (st3 = '')or (st4 = '') then exit;
         dDate:=StrToDate(USDateToStd(st4));
         if ((dDate=0) or (dDate=iDate1900) or (dDate=iDate2099)) then exit;
            //  nat + date+souche+N° +indice
         znum :=St+';'+FormatDateTime('dd/mm/yyyy',DDate)+';'+ St2 +';' +St3+';'+'0;;';
         AppelPiece([ znum , 'ACTION=CONSULTATION'],2);
         end
        else begin  // sinon double lcic simple,accès activite
         St := ReadTokenSt(Str);  // Affaire
         St2 := ReadTokenSt(Str); // Tiers
         St3 := ReadTokenSt(Str); // Date
         if (St = '') or (St2 = '') or (st3 = '') then exit;
         dDate:=StrToDate(USDateToStd(st3));
         if ((dDate=0) or (dDate=iDate1900) or (dDate=iDate2099)) then exit;
         AFCreerActiviteModale(tsaClient, tacGlobal, 'REA', '', St, St2, dDate  );
         end;
        END;
  705:  BEGIN      // activité  par affaire  (sur act)
        St := ReadTokenSt(Quoi);  // Affaire
        St2 := ReadTokenSt(Quoi); // Tiers
        St3 := ReadTokenSt(Quoi); // Dateactivite
        if (St = '') or (St2 = '') or (st3 = '') then exit;
        dDate:=StrToDate(USDateToStd(st3));
        if ((dDate=0) or (dDate=iDate1900) or (dDate=iDate2099)) then exit;
        AFCreerActiviteModale(tsaClient, tacGlobal, 'REA', '', St, St2, dDate  );
        END;
{$IFDEF BTP}
  750..751:  BEGIN      // MODIFBTP pour Saut de page
        FillChar(CleDoc,Sizeof(CleDoc),#0) ;
        AFZoomStringToCleDoc(Quoi,CleDoc) ;
        if Qui = 750 then
          // modif BRL temporaire pour décisionnel stock SNTE
          if CleDoc.naturepiece = 'CBT' then
            BtMajTenueStock (CleDoc)
          else
            BtMajSDP (CleDoc, (TheImpressionViaTOB <> nil))
        else
          BtMajSDP (CleDoc, True);
        RelancerEtat:=True;
        END;
{$ENDIF}
  END ;
Screen.Cursor:=crDefault ;

END ;

Procedure AFZoomStringToCleDoc ( StA : String ; Var CleDoc : R_CleDoc ) ;
Var StC : String ;
Begin
StC:=AnsiUppercase(Trim(ReadTokenSt(StA))) ; if StC='' then Exit ;
CleDoc.NaturePiece:=StC ;
StC:=AnsiUppercase(Trim(ReadTokenSt(StA)))  ; if StC='' then Exit else CleDoc.DatePiece := UsDateTimeToDateTime(StC);
StC:=AnsiUppercase(Trim(ReadTokenSt (StA))) ; if StC='' then Exit else CleDoc.Souche:=StC ;
StC:=AnsiUppercase(Trim(ReadTokenSt(StA)))  ; if StC='' then Exit else CleDoc.NumeroPiece:=StrtoInt(StC) ;
StC:=AnsiUppercase(Trim(ReadTokenSt(StA)))  ; CleDoc.Indice:=0 ;
{$IFDEF BTP}
StC:=AnsiUppercase(Trim(ReadTokenSt(StA)))  ; if StC='' then Exit else CleDoc.NumLigne:=StrtoInt(StC) ;
StC:=AnsiUppercase(Trim(ReadTokenSt(StA)))  ; if StC='' then Exit else CleDoc.NumOrdre:=StrtoInt(StC) ;
{$ENDIF}
end;


end.
