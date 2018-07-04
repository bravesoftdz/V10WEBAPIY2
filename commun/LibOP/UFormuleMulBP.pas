unit UFormuleMulBP;

interface

uses Sysutils,
    {$IFNDEF EAGLCLIENT}
    DB,
    {$ELSE}
    UTob,
    {$ENDIF}
    BPFctSession,
    BPBasic,
    HCtrls,
    BPUtil;

function FormuleMul(Func, Params, WhereSQL: hString; TT: TDataset; Total: Boolean):hString;
function FormuleMulOBJPREV(Func, Params, WhereSQL: hString; TT: TDataset; Total: Boolean; var S: hString):boolean;

implementation
function FormuleMulOBJPREV(Func, Params, WhereSQL: hString; TT: TDataset; Total: Boolean; var S: hString):boolean;
var Func2:hString;
    date:TDateTime;
    RI:integer;
begin
 Func2 := UpperCase(Func);
 Result:=true;
 S:='';
 //pour la fiche session coche valideé
 if Func2 = 'SESSIONVALIDE'
  then
   begin
    date:=StrToDateTime(TT.FindField('QBS_DATEVALIDATION').AsString);
    if date>10 then
    begin
      if ContextBP in [0,2,3] then
        S:='#ICO#' + IntToStr(47)
      else
        S:='#ICO#' + IntToStr(99);
    end;
   end
   //pour la fiche session coche eclat par delai
  else if Func2 = 'SESSIONECLATD'
  then
   begin
    if TT.FindField('QBS_SESSIONECLAT').AsString='X'
     then S:='#ICO#' + IntToStr(3);
   end
     //pour la fiche session coche eclat par taille
  else if Func2 = 'SESSIONECLATT'
  then
   begin
    if TT.FindField('QBS_SESSIONECLATAI').AsString='X'
     then S:='#ICO#' + IntToStr(12);
   end
  //pour la fiche session coche initialisée
  else if Func2 = 'SESSIONINIT'
  then
   begin
    if TT.FindField('QBS_SESSIONINIT').AsString='X'
     then S:='#ICO#' + IntToStr(46);
   end
  //pour la fiche session coche initialisée par délai
  else if Func2 = 'SESSIONINITD'
  then
   begin
    if TT.FindField('QBS_INITDELAI').AsString='X'
     then S:='#ICO#' + IntToStr(5);
   end
  //pour la fiche session coche initialisée coefficient
  else if Func2 = 'SESSIONINITCOEFF'
  then
   begin
    if TT.FindField('QBS_INITCOEFF').AsString='X'
     then S:='#ICO#' + IntToStr(47);
   end
  //pour la fiche session coche initialisée coefficient
  else if Func2 = 'SESSIONINITPREV'
  then
   begin
    if TT.FindField('QBS_INITPREVISION').AsString='X'
     then S:='#ICO#' + IntToStr(45);
   end
  //pour la fiche session statut initialisation
  else if (Func2 = 'SESSIONSTATUT') then
  begin
    if TT.FindField('QBS_INITDELAI').AsString='X' then
      S:='#ICO#' + IntToStr(77)
    else if TT.FindField('QBS_SESSIONECLAT').AsString='X' then
      S:='#ICO#' + IntToStr(24)
    else if TT.FindField('QBS_SESSIONINIT').AsString='X' then
      S:='#ICO#' + IntToStr(21);
  end
  //pour la fiche arbre
  else if Func2 = 'PRCTHISTOOBJ'
  then
   begin
    if BPOkOrli
     then
      //-----------------> ORLI
      begin
       if TT.FindField('QBR_REF1').AsFloat<>0
        then S:=FloatToStr((TT.FindField('QBR_OP1').AsFloat/
                                 TT.FindField('QBR_REF1').AsFloat)*100);
      end
      //ORLI <-----------------
     else
      begin
       RI:=DonneValeurAffiche(TT.FindField('QBR_CODESESSION').AsString);
       S:='100';
       case RI of
         1 : begin
              if TT.FindField('QBR_REF1').AsFloat<>0
               then S:=FloatToStr((TT.FindField('QBR_OP1').AsFloat/
                                        TT.FindField('QBR_REF1').AsFloat)*100);
             end;
         2 : begin
              if TT.FindField('QBR_QTEREF').AsFloat<>0
               then S:=FloatToStr((TT.FindField('QBR_QTEC').AsFloat/
                                 TT.FindField('QBR_QTEREF').AsFloat)*100);
             end;
         3 : begin
              if TT.FindField('QBR_REF2').AsFloat<>0
               then S:=FloatToStr((TT.FindField('QBR_OP2').AsFloat/
                                 TT.FindField('QBR_REF2').AsFloat)*100);
             end;
         4 : begin
              if TT.FindField('QBR_REF3').AsFloat<>0
               then S:=FloatToStr((TT.FindField('QBR_OP3').AsFloat/
                                 TT.FindField('QBR_REF3').AsFloat)*100);
             end;
         5 : begin
              if TT.FindField('QBR_REF4').AsFloat<>0
               then S:=FloatToStr((TT.FindField('QBR_OP4').AsFloat/
                                 TT.FindField('QBR_REF4').AsFloat)*100);
             end;
         6 : begin
              if TT.FindField('QBR_REF5').AsFloat<>0
               then S:=FloatToStr((TT.FindField('QBR_OP5').AsFloat/
                                 TT.FindField('QBR_REF5').AsFloat)*100);
             end;
         7 : begin
              if TT.FindField('QBR_REF6').AsFloat<>0
               then S:=FloatToStr((TT.FindField('QBR_OP6').AsFloat/
                                 TT.FindField('QBR_REF6').AsFloat)*100);
             end;
        end;
       end
       end
  //pour la fiche arbre
  else if Func2 = 'PRCTVARIATION'
  then
   begin
    RI:=DonneValeurAffiche(TT.FindField('QBR_CODESESSION').AsString);
    S:='100';
    case RI of
      1 : begin
           if TT.FindField('QBR_REF1').AsFloat<>0
            then S:=FloatToStr((TT.FindField('QBR_OP1').AsFloat/
                                     TT.FindField('QBR_REF1').AsFloat)*100);
          end;
      2 : begin
           if TT.FindField('QBR_QTEREF').AsFloat<>0
            then S:=FloatToStr((TT.FindField('QBR_QTEC').AsFloat/
                              TT.FindField('QBR_QTEREF').AsFloat)*100);
          end;
      3 : begin
           if TT.FindField('QBR_REF2').AsFloat<>0
            then S:=FloatToStr((TT.FindField('QBR_OP2').AsFloat/
                              TT.FindField('QBR_REF2').AsFloat)*100);
          end;
      4 : begin
           if TT.FindField('QBR_REF3').AsFloat<>0
            then S:=FloatToStr((TT.FindField('QBR_OP3').AsFloat/
                              TT.FindField('QBR_REF3').AsFloat)*100);
          end;
      5 : begin
           if TT.FindField('QBR_REF4').AsFloat<>0
            then S:=FloatToStr((TT.FindField('QBR_OP4').AsFloat/
                              TT.FindField('QBR_REF4').AsFloat)*100);
          end;
      6 : begin
           if TT.FindField('QBR_REF5').AsFloat<>0
            then S:=FloatToStr((TT.FindField('QBR_OP5').AsFloat/
                              TT.FindField('QBR_REF5').AsFloat)*100);
          end;
      7 : begin
           if TT.FindField('QBR_REF6').AsFloat<>0
            then S:=FloatToStr((TT.FindField('QBR_OP6').AsFloat/
                              TT.FindField('QBR_REF6').AsFloat)*100);
          end;
     end;
   end
  else
    Result := False;
end;

function FormuleMul(Func, Params, WhereSQL: hString; TT: TDataset; Total: Boolean):hString;
var S: hString;
begin
  if FormuleMulOBJPREV(Func, Params, WhereSQL,TT,Total,S)
    then Result:= S
    else Result:= DefProcCalcMul(Func,Params,WhereSQL,TT,Total);
end;

end.
