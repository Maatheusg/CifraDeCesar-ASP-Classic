<%@LANGUAGE="VBSCRIPT" LCID=1046 CODEPAGE="65001"%>
<%Option Explicit%>
<%
    dim i, letraUm, ind, temp, mensagem, msgCriptog, j, numDeslocamento, modo
    dim alfabeto(25)
    dim alfabetoOrganizado, compara,  maiuscula
    alfabeto(0) = "a"
    alfabeto(1) = "b"
    alfabeto(2) = "c"
    alfabeto(3) = "d"
    alfabeto(4) = "e"
    alfabeto(5) = "f"
    alfabeto(6) = "g"
    alfabeto(7) = "h"
    alfabeto(8) = "i"
    alfabeto(9) = "j"
    alfabeto(10) = "k"
    alfabeto(11) = "l"
    alfabeto(12) = "m"
    alfabeto(13) = "n"
    alfabeto(14) = "o"
    alfabeto(15) = "p"
    alfabeto(16) = "q"
    alfabeto(17) = "r"
    alfabeto(18) = "s"
    alfabeto(19) = "t"
    alfabeto(20) = "u"
    alfabeto(21) = "v"
    alfabeto(22) = "w"
    alfabeto(23) = "x"
    alfabeto(24) = "y"
    alfabeto(25) = "z"

    'A função percorre todo o alfabeto reorganizando baseado no deslocamento enviado'
    'O alfabeto organizado é enviado para um outro vetor'
    sub organizaAlfabeto(deslocamento)
        redim alfabetoOrganizado(25)
        for i = 0 to ubound(alfabeto)
            letraUm = alfabeto(i)
            ind = i + deslocamento
            if ind > 25 then
                temp = ind - 25
                ind = 0
               ind = ind + temp - 1
            end if
            alfabetoOrganizado(ind) = letraUm
        next
    end sub

    function criptografar(mensagem)
        mensagem = trim(mensagem)
        redim msgCriptog(len(mensagem))
        'Percorre toda a mensagem, pega uma única letra'
        'Chama a função que acha o índice da letra no alfabeto original'
        'O novo vetor recebe a letra representada no índice'
        for i = 0 to len(mensagem) - 1
            compara = mid(mensagem, i + 1, 1)
            call letrasComAcento(compara)
                if modo = "1" then 'Modo 1 = Criptografar  |  Modo 2 = Descriptografar'
                    temp = acharIndice(compara)
                    if temp = -1 then
                        msgCriptog(i) = compara
                    elseif maiuscula then
                        msgCriptog(i) = ucase(alfabetoOrganizado(temp))
                    else
                        msgCriptog(i) = alfabetoOrganizado(temp)
                    end if                    
                else
                    temp = acharIndiceDesc(compara)
                    if temp = -1 then
                        msgCriptog(i) = compara
                    elseif maiuscula then
                        msgCriptog(i) = ucase(alfabeto(temp))
                    else
                        msgCriptog(i) = alfabeto(temp)
                    end if
                end if
        next
        criptografar = msgCriptog
    end function

    function acharIndice(letra)
        maiuscula = false 'sinaliza se a letra é maiúscula ou não
        'Percorre o alfabeto em busca da letra e retorna o índice dela'
        for j = 0 to ubound(alfabeto)
            if letra = alfabeto(j) then
                acharIndice = j
                exit function
                
            elseif letra = ucase(alfabeto(j)) then
                maiuscula = true
                acharIndice = j
                exit function
            end if
        next
        acharIndice = -1
    end function

    'Função igual a anterior, mas os alfabetos comparados são invertidos'
    function acharIndiceDesc(letra)
        maiuscula = false
        for j = 0 to ubound(alfabetoOrganizado)
            if letra = alfabetoOrganizado(j) then
                acharIndiceDesc = j
                exit function
            elseif letra = ucase(alfabetoOrganizado(j)) then
                maiuscula = true
                acharIndiceDesc = j
                exit function
            end if
        next
        acharIndiceDesc = -1
    end function

    sub letrasComAcento(letra)
        if letra = "á" or letra = "â" or letra = "ã" or letra = "à" then
            letra = "a"
        elseif letra = "é" or letra = "ê" or letra = "è" then
            letra = "e"
        elseif letra = "í" or letra = "î" or letra = "ì" then
            letra = "i"
        elseif letra = "ó" or letra = "ô" or letra = "õ" or letra = "ò" then
            letra = "o"
        elseif letra = "ú" or letra = "û" or letra = "ù" or letra = "ü" then
            letra = "u"   
        elseif letra = "ç" then
            letra = "c" 
        end if
    end sub   

    sub mostraTexto(text)
        for i = 0 to ubound(text)            
            response.write(text(i))        
        next
    end sub

    numDeslocamento = request.form("numDeslocamento")
    mensagem = request.form("texto")
    modo = request.form("desOuCrip")
    call organizaAlfabeto(cint(numDeslocamento))
    msgCriptog = criptografar(mensagem)

%>

    <html>
        <head>
            <meta charset="utf-8">
            <title>Cifra de César</title>
            <link rel="preconnect" href="https://fonts.googleapis.com">
            <link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
            <link href="https://fonts.googleapis.com/css2?family=Assistant&family=Bebas+Neue&display=swap" rel="stylesheet">
            <link rel="stylesheet" href="style.css">
        </head>
        <body>
            <h1>Cifra de César</h1>
            <h2>Deslocamento</h2>
            <form method="post" action="index.asp">
                <input id="inputNumero" type="number" name="numDeslocamento" placeholder="Insira um número" min="1" max="25" required>
                <div>
                    <label><input name="desOuCrip" value="1" class="opcao" type="radio" checked>Criptografar</label>
                    <br>
                    <label><input name="desOuCrip" value="2" class="opcao" type="radio">Descriptografar</label>        
                    <h2>Mensagem</h2>
                    <textarea name="texto" id="areaTexto" placeholder="Insira a mensagem" cols="100" rows="15" required><%call mostraTexto(msgCriptog)%></textarea>
                </div>
                <input id="btnEnviar" type="submit">
            </form>
        </body>
    </html>