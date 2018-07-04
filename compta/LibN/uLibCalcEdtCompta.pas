unit uLibCalcEdtCompta;

interface

uses uTob; // Tob

var FTobCumulAnoGL   : Tob;
    FTobCumulAuGL    : Tob;
    FTobBilanGestion : Tob;
    FTobModePaie     : Tob;

// Fonction d'interface pour le générateur d'état
function CalcOLEEtat(sf, sp : string) : variant;
function ReadParam(var St : String) : string;

// Fonctions appelées par les fonctions @
function LibelleSousSection(Axe, Niveau, Code : string) : string;
function GetRepart(Axe, Cle : string) : string;
function GetCumulGrandLivre( vStGL, vStGLPar, vStAno, vStData1, vStData2, vStChamp : string) : variant;
function SetCumulBilanGestion( vStCompte : string; vDebit, vCredit : Double) : variant;
function GetCumulBilanGestion( vStNature, vStChamp : string) : variant;

function SetCumulGLAuxi( vDebit, vCredit : Double) : variant;
function GetCumulGLAuxi( vStChamp : string) : variant;

function GetTobModePaie(Where : String) : String;

implementation

uses
{$IFDEF EAGLCLIENT}
{$ELSE}
   DB,
  {$IFNDEF DBXPRESS}dbtables{$ELSE}uDbxDataSet{$ENDIF},
{$ENDIF}

{$IFDEF MODENT1}
   CPTypeCons,
{$ENDIF MODENT1}
   HCtrls,        // OpenSQL
   HEnt1,         // IsNumeric
   Variants,      // NULL
   CalcOle,       //
   ParamSoc,      // SetParamSoc
   SysUtils,      // StrToDateTime
   Ent1,          // TExodate
   uLibExercice;  // WhatDate

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... :
Créé le ...... :
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... :
*****************************************************************}
function CalcOLEEtat(sf,sp : string) : variant ;
var s1,s2,s3,s4,s5,s6,s7,s8,s9,s10 : string ;
    d1,d2 : TDatetime ;
    i1 : Integer ;
    Exo : TExoDate ;
    vTob : TOB;
begin
  if sf='LIBELLESOUSSECTION' then
     begin
     s1:=ReadParam(sp) ;
     s2:=ReadParam(sp) ;
     s3:=ReadParam(sp) ;
     result:=LibelleSousSection(s1,s2,s3) ;
     exit ;
     end else
  if sf='DATECUMUL' then
     begin
     s1:=ReadParam(sp) ;
     s2:=ReadParam(sp) ;
     result:=Get_DateCumul(s1,s2) ;
     exit ;
     end else
  if sf='GETCUMUL' then
     begin
     s1:=ReadParam(sp) ;
     s2:=ReadParam(sp) ;
     s3:=ReadParam(sp) ;
     s4:=ReadParam(sp) ;
     s5:=ReadParam(sp) ;
     s6:=ReadParam(sp) ;
     s7:=ReadParam(sp) ;
     result:=Get_cumul(s1,s2,s3,s4,s5,s6,s7) ;
     exit ;
     end else
  if sf='GETCUMUL2' then
     begin
     s1:=ReadParam(sp) ;
     s2:=ReadParam(sp) ;
     s3:=ReadParam(sp) ;
     s4:=ReadParam(sp) ;
     s5:=ReadParam(sp) ;
     s6:=ReadParam(sp) ;
     s7:=ReadParam(sp) ;
     s8:=ReadParam(sp) ;
     s9:=ReadParam(sp) ;
     s10:=ReadParam(sp) ;
     result:=Get_cumul2(s1,s2,s3,s4,s5,s6,s7,s8,s9,s10,'',NULL) ;
     exit ;
     end else
  if sf='GETCUMULDOUBLE' then
     begin
     s1:=ReadParam(sp) ;
     s2:=ReadParam(sp) ;
     s3:=ReadParam(sp) ;
     s4:=ReadParam(sp) ;
     s5:=ReadParam(sp) ;
     s6:=ReadParam(sp) ;
     s7:=ReadParam(sp) ;
     s8:=ReadParam(sp) ;
     result:=Get_CumulDouble(s1,s2,s3,s4,s5,s6,s7,s8) ;
     exit ;
     end else
  if sf='GETCONSTANTE' then
     begin
     s1:=ReadParam(sp) ;
     s2:=ReadParam(sp) ;
     s3:=ReadParam(sp) ;
     result:=Get_constante(s1,s2,s3) ;
     exit ;
     end else
  if sf='GETINFO' then
     begin
     s1:=ReadParam(sp) ;
     s2:=ReadParam(sp) ;
     s3:=ReadParam(sp) ;
     s4:=ReadParam(sp) ;
     result:=Get_info(s1,s2,s3,s4) ;
     exit ;
     end else
  if sf='GETDATEDEBUT' then
     begin
     s1:=ReadParam(sp) ;
     WhatDate(s1,d1,d2,i1,exo) ;
     result := d1 ;
     exit ;
     end else
  if sf='GETDATEFIN' then
     begin
     s1:=ReadParam(sp) ;
     WhatDate(s1,d1,d2,i1,exo) ;
     result := d2 ;
     exit ;
     end else
  {$IFDEF GCGC}
  if sf='GETPIEDBASE' then
     begin
     s1:=ReadParam(sp) ;
     s2:=ReadParam(sp) ;
     s3:=ReadParam(sp) ;
     s4:=ReadParam(sp) ;
     s5:=ReadParam(sp) ;
     s6:=ReadParam(sp) ;
     s7:=ReadParam(sp) ;
     d1:=DateOnly(StrToDateTime(s5)) ;
     i1:=StrToInt(s6) ;
  //   result := Get_PiedBase(s1,s2,s3,s4,d1,i1,s7) ;
     exit ;
     end else
  if sf='GETPIEDECHE' then
     begin
     s1:=ReadParam(sp) ;
     s2:=ReadParam(sp) ;
     s3:=ReadParam(sp) ;
     s4:=ReadParam(sp) ;
     s5:=ReadParam(sp) ;
     s6:=ReadParam(sp) ;
     s7:=ReadParam(sp) ;
     d1:=DateOnly(StrToDateTime(s5)) ;
     i1:=StrToInt(s6) ;
  //   result := Get_PiedEche(s1,s2,s3,s4,d1,i1,s7) ;
     exit ;
     end else
  {$ENDIF}
  {--------------------------------------------------------------
  Fonction CP_GetCumulCpteJal : 7 paramètres
  Param 1 : Racine de Compte
  Param 2 : Journal
  Param 3 : Date début (0 signifie V_PGI.DateEntree)
  Param 4 : Date fin (0 signifie V_PGI.DateEntree)
  Param 5 : Type de valeur (TC : Total crédit, TD : Total débit)
  Param 6 : Débit négatif : NEG sinon rien
  Param 7 : A nouveaux (T : tous)
  ---------------------------------------------------------------}
  if sf='CP_GETCUMULCPTEJAL' then
     begin
     s1:=ReadParam(sp) ;
     s2:=ReadParam(sp) ;
     d1:=DateOnly(StrToDateTime(ReadParam(sp))) ;
     d2:=DateOnly(StrToDateTime(ReadParam(sp))) ;
     s3:=ReadParam(sp) ;
     s4:=ReadParam(sp) ;
     s5:=ReadParam(sp) ;
     result:=CPGetCumulCpteJal(s1,s2,d1,d2,s3,s4,s5) ;
     exit ;
     end else
  {--------------------------------------------------------------
  Fonction CP_GetCumulCorresp : 7 paramètres
  Param 1 : Compte de correspondance
  Param 2 : Plan 1 ou 2
  Param 3 : Etablissement
  Param 4 : Devise
  Param 5 : Période (idem GetCumul)
  Param 6 : Type de valeur (TC : Total crédit, TD : Total débit)
  Param 7 : Débit négatif : NEG sinon rien
  ---------------------------------------------------------------}
  if sf='CP_GETCUMULCORRESP' then
     begin
     s1:=ReadParam(sp) ;
     s2:=ReadParam(sp) ;
     s3:=ReadParam(sp) ;
     s4:=ReadParam(sp) ;
     s5:=ReadParam(sp) ;
     s6:=ReadParam(sp) ;
     s7:=ReadParam(sp) ;
     result:=CPGetCumulCorresp(s1,s2,s3,s4,s5,s6,s7) ;
     exit ;
     end else
  if sf='GETCUMULDC' then
     begin
     s1:=ReadParam(sp) ; //TypePlan
     s2:=ReadParam(sp) ; //Compte
     s3:=ReadParam(sp) ; //Exo
     d1:=DateOnly(StrToDateTime(ReadParam(sp))) ; //Date1
     d2:=DateOnly(StrToDateTime(ReadParam(sp))) ; //Date2
     s6:=ReadParam(sp) ; //Type résultat
     result:=GetCumulDC(s1,s2,s3,d1,d2,s6) ;
     exit ;
     end else
  if sf='GETCUMULBAL' then
     begin
     s1:=ReadParam(sp) ;
     s2:=ReadParam(sp) ;
     s3:=ReadParam(sp) ;
     s4:=ReadParam(sp) ;
     s5:=ReadParam(sp) ;
     s6:=ReadParam(sp) ;
     s7:=ReadParam(sp) ;
     s8:=ReadParam(sp) ;
     s9:=ReadParam(sp) ;
     s10:=ReadParam(sp) ;
     result:=Get_CumulBal(s1,s2,s3,s4,s5,s6,s7,s8,s9,s10) ;
     Exit ;
     end else
      if sf='GETONLYNUMBER' then
      begin
          s1:=ReadParam(sp) ;
          for i1 := Length(s1) downto 1 do if (not (s1[i1] in ['0'..'9'])) then Delete(s1, i1,1);
          Result := s1;
      end
      else
  // Fin BPY
  if sf='ADDAFTER' then begin
    s1:=ReadParam(sp) ;
    s2:=ReadParam(sp) ;
    if IsNumeric(s1) then
      SetParamSoc(s2, IntToStr(StrToInt(s1)+1));
    Result := S1;
  end else
  {$IFNDEF EAGLCLIENT}
  {$IFNDEF EAGLSERVER}
  if sf='INFODOSSIERPIED' then
  begin
    if (ctxPCL in V_PGI.PGIContexte) then
      Result := V_PGI.DefaultSectionName+ ' / '+TraduireMemoire('Dossier n°')+V_PGI.NoDossier
    else
      Result := GetParamSocSecur('SO_LIBELLE', '');
    end
    else
  {$ENDIF}
  {$ENDIF}

  if sf='GETTVATPFINFO' then begin
     s1:=ReadParam(sp) ;
     s2:=ReadParam(sp) ;
     s3:=ReadParam(sp) ;
     s4:=ReadParam(sp) ;
     s5:=ReadParam(sp) ;
     s6:=ReadParam(sp) ;
     result:=Get_TVATPFINFO(s1,s2,s3,s4,s5,s6) ;
  end else
  if sf='GETINFORIB' then begin
     s1:=ReadParam(sp) ; //Le RIB
     s2:=ReadParam(sp) ; //Quelle Info
     Result:=Get_RIBINFO(S1,S2) ;
  end else

  if sf='GETREPART' then begin // Utilisé dans l'état CST;ANA
     s1:=ReadParam(sp) ; // Axe
     s2:=ReadParam(sp) ; // Clé
     Result := GetRepart(S1,S2);
  end else
  if sf='GETSOUSSECTIONS' then begin // Utilisé dans l'état CST;ANA + CST;AN2
     s1:=ReadParam(sp) ; // Axe
     s2:=ReadParam(sp) ; // Section
     Result := GetSousSections(S1,S2);
  end else

  // GCO - 29/11/2005
  if sf='GETCUMULGRANDLIVRE' then
  begin
    s1 := ReadParam(sp); // GLGENE ou GLAUXI
    s2 := ReadParam(sp); // Si GLGENEPARAUXI ou GLAUXIPARGENE ==> "X" sinon "-"
    s3 := ReadParam(sp); // "X" pour ANO, "-" pour CUMULAU
    s4 := ReadParam(sp); // E_GENERAL ou E_AUXILIAIRE
    s5 := ReadParam(sp); // E_AUXILIAIRE ou E_GENERAL
    s6 := ReadParam(sp); // Champ
    Result := GetCumulGrandLivre(s1, s2, s3, s4, s5, s6);
  end else

  if sf='SETCUMULBILANGESTION' then
  begin
    s1:=ReadParam(sp); // General
    s2:=ReadParam(sp); // Débit
    s3:=ReadParam(sp); // Crédit
  //  Result := SetCumulBilanGestion(S1, StrToFloat(S2), StrToFloat(S3));
    Result := SetCumulBilanGestion(S1, Valeur(S2), Valeur(S3));
  end
  else
  if sf='GETCUMULBILANGESTION' then
  begin
    s1:=ReadParam(sp); // Nature (BIL, CHA, PRO, EXT)
    s2:=ReadParam(sp); // Champ  (DEBIT, CREDIT)
    Result := GetCumulBilanGestion(S1, S2);
  end
  else
  if sf='SETCUMULGLAUXI' then
  begin
    s1:=ReadParam(sp); // Débit
    s2:=ReadParam(sp); // Crédit
  //  Result := SetCumulGLAuxi(StrToFloat(S1), StrToFloat(S2));
    Result := SetCumulGLAuxi(Valeur(S1), Valeur(S2));
  end
  else
  if sf='GETCUMULGLAUXI' then
  begin
    s1:=ReadParam(sp); // Débit ou Crédit
    Result := GetCumulGLAuxi(S1);
  end
  // BVE 29.08.07
  // Permet de recuperer le nième mois apres la date passée en parametre
  // p1 : La date de depart
  // p2 : Le nombre de mois apres la date de départ.
  else
  if sf='GETMOIS' then
  begin
    s1:=ReadParam(sp); // Date de départ
    s2:=ReadParam(sp); // Nb Mois
    Result := PlusMois(StrToInt(s1),StrToInt(s2));
  end
  // BVE 30.08.07
  // Pour l'Etat CLJ/CIV FQ 21313
  else
  if sf='INSERTEDTLEGAL' then
  begin
    s1 := ReadParam(sp); // Les valeurs à insérer
    Result := ExecuteSQL('INSERT INTO EDTLEGAL ' +
    '(ED_OBLIGATOIRE, ED_TYPEEDITION, ED_EXERCICE, ED_DATE1, ED_DATE2, ED_DATEEDITION, ED_NUMEROEDITION, ED_CUMULDEBIT, ED_CUMULCREDIT, ED_CUMULSOLDED, ED_CUMULSOLDEC, ED_UTILISATEUR, ED_DESTINATION) ' +
    'VALUES (' + s1 + ')');
  end
  else
  // BVE 30.08.07
  // Pour l'Etat CLJ/CIV FQ 21313
  if sf='UPDATEEDTLEGAL' then
  begin
    s1 := ReadParam(sp); // Les valeurs à insérer
    Result := ExecuteSQL('UPDATE EDTLEGAL SET ' + s1 );
  end
  else
  // FQ 21806
  // BVE 08.11.07
  // Etat ECH/ECH probleme de lenteur
  if sf='GETMODEPAIE' then
  begin
    s1 := ReadParam(sp); // Where
    s2 := ReadParam(sp); // Mode Paie
    if not(assigned(FTobModePaie)) then GetTobModePaie(s1);
    vTob := FTobModePaie.FindFirst(['E_MODEPAIE'],[s2],true);
    if vTob <> nil then
       Result := vTob.GetString('SOMME');
  end
  else  if sf='DELETEMODEPAIE' then
  begin
    if assigned(FTobModePaie) then FreeAndNil(FTobModePaie);
  end
  else  ;
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 15/11/2007
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
function ReadParam(var St : String) : string ;
var i : Integer ;
begin
  i:=Pos(';',St) ;
  if i<=0 then
    i:=Length(St)+1 ;

  Result:=Copy(St,1,i-1) ;
  Delete(St,1,i) ;
end ;

////////////////////////////////////////////////////////////////////////////////
////////////////////////////////////////////////////////////////////////////////
////////////////////////////////////////////////////////////////////////////////
////////////////////////////////////////////////////////////////////////////////
////////////////////////////////////////////////////////////////////////////////
////////////////////////////////////////////////////////////////////////////////
////////////////////////////////////////////////////////////////////////////////
////////////////////////////////////////////////////////////////////////////////
////////////////////////////////////////////////////////////////////////////////
////////////////////////////////////////////////////////////////////////////////

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... :
Créé le ...... :   /  /
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
function LibelleSousSection(Axe, Niveau, Code : string) : string;
var Q : TQuery ;
    s : String ;
begin
  s := 'SELECT * FROM SSSTRUCR WHERE PS_AXE="' + Axe + '" AND PS_SOUSSECTION="' + Niveau + '"' ;
  s := s + ' AND PS_CODE="' + Code + '"' ;
  Q := OpenSQL(s,True) ;
  if not Q.EOF then
    Result := Q.FindField('PS_LIBELLE').AsString
  else
    Result:= '';
  Ferme(Q) ;
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... :
Créé le ...... :
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
function GetRepart(Axe, Cle : string) : string;
var Q  : TQuery;
begin
  Q := OpenSQL('SELECT V_SECTION FROM VENTIL WHERE V_NATURE="CL'+Copy(Axe,2,1)+'" AND V_COMPTE="'+Cle+'"', True);
  While Not Q.EOF do
  begin
    Result := Result  + Q.Fields[0].AsString + ' ; ';
    Q.Next;
  end;
  Ferme(Q);
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 29/11/2005
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
function GetCumulGrandLivre( vStGL, vStGLPar, vStAno, vStData1, vStData2, vStchamp : string) : Variant;
var lTobTrouve : Tob;
begin
  Result := 0;
  lTobTrouve := nil;

  if (FTobCumulAnoGL = nil) or (FTobCumulAuGL = nil) then Exit;

  // E_GENERAL toujours dans DATA1
  // E_AUXILIAIRE toujours dans DATA2
  if vStGL = 'GLGENE' then
  begin
    if vStGLPar = 'X' then
    begin // Grand livre général par auxiliaire
      if vStAno = 'X' then
        lTobTrouve := FTobCumulAnoGL.FindFirst(['E_GENERAL','E_AUXILIAIRE'],[vStData1,vStData2], False)
      else
        lTobTrouve := FTobCumulAuGL.FindFirst(['E_GENERAL','E_AUXILIAIRE'],[vStData1,vStData2], False);
    end
    else
    begin // Grand livre général
      if vStAno = 'X' then
        lTobTrouve := FTobCumulAnoGL.FindFirst(['E_GENERAL'],[vStData1], False)
      else
        lTobTrouve := FTobCumulAuGL.FindFirst(['E_GENERAL'],[vStData1], False);
    end;
  end;

  if vStGL = 'GLAUXI' then
  begin
    if vStGLPar = 'X' then
    begin // Grand livre auxiliaire par général
      if vStAno = 'X' then
        lTobTrouve := FTobCumulAnoGL.FindFirst(['E_AUXILIAIRE','E_GENERAL'],[vStData2, vStData1], False)
      else
        // GCO - 13/11/2007 - FQ 21841 inversion de data2 et data1
        //lTobTrouve := FTobCumulAuGL.FindFirst(['E_AUXILIAIRE','E_GENERAL'],[vStData1,vStData2], False);
        lTobTrouve := FTobCumulAuGL.FindFirst(['E_AUXILIAIRE','E_GENERAL'],[vStData2,vStData1], False);
    end
    else
    begin // Grand livre auxiliaire
      if vStAno = 'X' then
        lTobTrouve := FTobCumulAnoGL.FindFirst(['E_AUXILIAIRE'],[vStData2], False)
      else
        lTobTrouve := FTobCumulAuGL.FindFirst(['E_AUXILIAIRE'],[vStData2], False);
    end;
  end;

  // Y_SECTION toujours dans DATA1
  // Y_GENERAL toujours dans DATA2
  if vStGL = 'GLANA' then
  begin
    if vStGLPar = 'X' then // Grand Analytique par Général ou Général par Analytique
    begin
      if vStAno = 'X' then
        lTobTrouve := FTobCumulAnoGL.FindFirst(['Y_SECTION','Y_GENERAL'],[vStData1,vStData2], False)
      else
        lTobTrouve := FTobCumulAuGL.FindFirst(['Y_SECTION','Y_GENERAL'],[vStData1,vStData2], False);
    end
    else
    begin
      if vStAno = 'X' then // Grand Analytique
        lTobTrouve := FTobCumulAnoGL.FindFirst(['Y_SECTION'],[vStData1], False)
      else
        lTobTrouve := FTobCumulAuGL.FindFirst(['Y_SECTION'],[vStData1], False);
    end;
  end;

  if lTobTrouve <> nil then
    Result := lTobTrouve.GetDouble(vStChamp);
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 09/12/2005
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
function SetCumulBilanGestion( vStCompte : string; vDebit, vCredit : Double) : variant;
var lTobTrouve : Tob;
begin
{$IFDEF NOVH}
Result := '';
{$ELSE}
  // Test Compte Bilan
  if ( ( vStCompte >= VH^.FBil[1].Deb ) and ( vStCompte <= VH^.FBil[1].Fin ) ) or // 1er groupe Bilan
     ( ( vStCompte >= VH^.FBil[2].Deb ) and ( vStCompte <= VH^.FBil[2].Fin ) ) or // 2ème groupe Bilan
     ( ( vStCompte >= VH^.FBil[3].Deb ) and ( vStCompte <= VH^.FBil[3].Fin ) ) or // 3ème groupe Bilan
     ( ( vStCompte >= VH^.FBil[4].Deb ) and ( vStCompte <= VH^.FBil[4].Fin ) ) or // 4ème groupe Bilan
     ( ( vStCompte >= VH^.FBil[5].Deb ) and ( vStCompte <= VH^.FBil[5].Fin ) ) or // 5ème groupe Bilan
     ( vStCompte = VH^.OuvreBil ) or ( vStCompte = GetParamSocSecur('SO_FERMEBIL', '') )   // Compte Bilan (ouverture / fermeture)
     then Result := 'BIL'
  // Test Compte Charge
  else
  if ( ( vStCompte >= VH^.FCha[1].Deb ) and ( vStCompte <= VH^.FCha[1].Fin ) ) or // 1er groupe Bilan
     ( ( vStCompte >= VH^.FCha[2].Deb ) and ( vStCompte <= VH^.FCha[2].Fin ) ) or // 2ème groupe Bilan
     ( ( vStCompte >= VH^.FCha[3].Deb ) and ( vStCompte <= VH^.FCha[3].Fin ) ) or // 3ème groupe Bilan
     ( ( vStCompte >= VH^.FCha[4].Deb ) and ( vStCompte <= VH^.FCha[4].Fin ) ) or // 4ème groupe Bilan
     ( ( vStCompte >= VH^.FCha[5].Deb ) and ( vStCompte <= VH^.FCha[5].Fin ) )    // 5ème groupe Bilan
     then Result := 'CHA'
  // Test Compte Produit
  else
  if ( ( vStCompte >= VH^.FPro[1].Deb ) and ( vStCompte <= VH^.FPro[1].Fin ) ) or // 1er groupe Bilan
     ( ( vStCompte >= VH^.FPro[2].Deb ) and ( vStCompte <= VH^.FPro[2].Fin ) ) or // 2ème groupe Bilan
     ( ( vStCompte >= VH^.FPro[3].Deb ) and ( vStCompte <= VH^.FPro[3].Fin ) ) or // 3ème groupe Bilan
     ( ( vStCompte >= VH^.FPro[4].Deb ) and ( vStCompte <= VH^.FPro[4].Fin ) ) or // 4ème groupe Bilan
     ( ( vStCompte >= VH^.FPro[5].Deb ) and ( vStCompte <= VH^.FPro[5].Fin ) )    // 5ème groupe Bilan
     then Result := 'PRO'
  // Sinon Extra-comptable
  else
    Result := 'EXT' ;

  if FTobBilanGestion <> nil then
  begin
    lTobTrouve := FTobBilanGestion.FindFirst(['NATURE'], [Result], False);
    if lTobTrouve <> nil then
    begin
      lTobTrouve.SetDouble('DEBIT', lTobTrouve.GetDouble('DEBIT') + vDebit);
      lTobTrouve.SetDouble('CREDIT', lTobTrouve.GetDouble('CREDIT') + vCredit);
    end;
  end;
{$ENDIF NOVH}
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 09/12/2005
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
function GetCumulBilanGestion( vStNature, vStChamp : string) : Variant;
var lTobTrouve : Tob;
begin
  Result := 0;
  if FTobBilanGestion <> nil then
  begin
    lTobTrouve := FTobBilanGestion.FindFirst(['NATURE'],[vStNature], False);
    if lTobTrouve <> nil then
      Result := lTobTrouve.GetDouble(vStChamp);
  end;
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 28/03/2006
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
function SetCumulGLAuxi( vDebit, vCredit : Double) : Variant;
var Debit, Credit : Double;
begin
  Result := 0;
  if FTobBilanGestion <> nil then
  begin
    Debit  := FTobBilanGestion.GetDouble('DEBIT');
    Credit := FTobBilanGestion.GetDouble('CREDIT');

    FTobBilanGestion.SetDouble('DEBIT',  Debit + vDebit);
    FTobBilanGestion.SetDouble('CREDIT', Credit + vCredit);
  end;
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 28/03/2006
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
function GetCumulGLAuxi(vStChamp : string) : Variant;
begin
  if FTobBilanGestion <> nil then
  begin
    Result := FTobBilanGestion.GetDouble(vStChamp);
  end;
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : BVE
Créé le ...... : 05/11/2007
Modifié le ... :   /  /
Description .. : FQ 21806 BVE 08.11.07
Mots clefs ... : 
*****************************************************************}
function GetTobModePaie(Where : String) : String;
var SQL : String;
begin
  SQL := 'SELECT E_MODEPAIE, COUNT(E_MODEPAIE) SOMME ' +
         'FROM ECRITURE ' +
         'LEFT JOIN TIERS ON E_AUXILIAIRE=T_AUXILIAIRE ' +
         'LEFT JOIN GENERAUX ON E_GENERAL=G_GENERAL ' +
         'LEFT JOIN MODEPAIE ON E_MODEPAIE=MP_MODEPAIE ' +
         Where +
         ' GROUP BY E_MODEPAIE';
  if FTobModePaie <> nil then FTobModePaie.ClearDetail
  else FTobModePaie := TOB.Create('Mode de paiement',nil,-1);
  FTobModePaie.LoadDetailDBFromSQL('ECRITURE',SQL);
end;

////////////////////////////////////////////////////////////////////////////////

end.
