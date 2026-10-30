"""Autocorreção de cargo, setor e empreendimento fora do padrão.

Para cada valor, em ordem:
  1. Igual à grafia oficial                                   -> ok
  2. Mesmo valor com outra caixa ou espaços                   -> autocorreção
  3. Está em config/apelidos.csv (decisão já tomada por você) -> autocorreção
  4. Igual a um oficial ignorando acentos, "(a)", forma
     feminina/masculina e abreviações (Tec., Aux., Coord....) -> autocorreção
  5. Parecido com um oficial (semelhança >= LIMIAR)           -> sugestão (só aplica se você aprovar)
  6. Nada parecido                                            -> sem correspondência (fica como pendência)

Os passos 2 a 4 não são palpites: são regras exatas ou decisões suas. O passo 5 é palpite,
por isso nunca é aplicado sozinho.
"""
import difflib
import re
import unicodedata

from core.referencias import Referencias, chave

LIMIAR_SUGESTAO = 0.80
CAMPOS = ("Cargo", "Setor", "Empreendimento")

ABREVIACOES = [
    (r"\btec\b\.?", "tecnico"), (r"\baux\b\.?", "auxiliar"), (r"\bcoord\b\.?", "coordenador"),
    (r"\benc\b\.?", "encarregado"), (r"\bdepto\b\.?", "departamento"), (r"\bdep\b\.?", "departamento"),
    (r"\bt\.\s*i\b\.?", "ti"), (r"\bjr\b\.?", "junior"), (r"\bsr\b\.?", "senior"),
    (r"\bmanut\b\.?", "manutencao"), (r"\bassist\b\.?", "assistente"), (r"\bcond\b\.?", "condominio"),
    (r"\bres\b\.?", "residencial"), (r"\beng\b\.?", "engenheiro"),
]
_PALAVRA_A = re.compile(r"([^\W\d_]+)\(a\)", re.UNICODE)


def forte(texto: str) -> str:
    """Chave tolerante: sem acentos, sem "(a)", com abreviações expandidas e espaços normalizados."""
    s = unicodedata.normalize("NFKD", str(texto or "")).encode("ascii", "ignore").decode().lower()
    s = re.sub(r"\s*\(\s*[ao]\s*\)", "", s)
    s = re.sub(r"\s*-\s*", " - ", s)
    for padrao, troca in ABREVIACOES:
        s = re.sub(padrao, troca, s)
    return " ".join(s.split())


def _feminino(palavra: str) -> str:
    if palavra.endswith("o"):
        return palavra[:-1] + "a"
    return palavra + "a"  # coordenador -> coordenadora, gestor -> gestora


def variantes(oficial: str) -> set[str]:
    """Formas aceitas de um valor oficial: ele mesmo, a masculina e a feminina."""
    formas = {oficial}
    if "(a)" in oficial:
        formas.add(_PALAVRA_A.sub(lambda m: m.group(1), oficial))
        formas.add(_PALAVRA_A.sub(lambda m: _feminino(m.group(1)), oficial))
    return formas


def _oficiais(campo: str, refs: Referencias) -> dict:
    return {"Cargo": refs.nomes_cargos, "Setor": refs.setores, "Empreendimento": refs.empreendimentos}[campo]


def _indice(campo: str, refs: Referencias) -> dict:
    """forte(variante) -> oficial; chaves ambíguas (dois oficiais) ficam fora."""
    cache = refs.__dict__.setdefault("_indice_autocorrecao", {})
    if campo not in cache:
        indice, ambiguas = {}, set()
        for oficial in _oficiais(campo, refs).values():
            for v in variantes(oficial):
                k = forte(v)
                if k in indice and indice[k] != oficial:
                    ambiguas.add(k)
                indice[k] = oficial
        cache[campo] = {k: v for k, v in indice.items() if k not in ambiguas}
    return cache[campo]


def classificar(campo: str, valor: str, refs: Referencias) -> tuple[str, str, float]:
    """Devolve (tipo, valor corrigido ou sugerido, semelhança).

    tipo: "vazio", "sem lista", "ok", "auto", "sugestao" ou "sem".
    """
    if not str(valor or "").strip():
        return "vazio", "", 0.0
    oficiais = _oficiais(campo, refs)
    if not oficiais:
        return "sem lista", valor, 0.0
    if valor in oficiais.values():
        return "ok", valor, 1.0
    if chave(valor) in oficiais:
        return "auto", oficiais[chave(valor)], 1.0
    apelido = refs.apelidos.get(campo, {}).get(chave(valor))
    if apelido and chave(apelido) in oficiais:
        return "auto", oficiais[chave(apelido)], 1.0
    indice = _indice(campo, refs)
    k = forte(valor)
    if k in indice:
        return "auto", indice[k], 1.0
    perto = difflib.get_close_matches(k, list(indice), n=1, cutoff=LIMIAR_SUGESTAO)
    if perto:
        return "sugestao", indice[perto[0]], round(difflib.SequenceMatcher(None, k, perto[0]).ratio(), 2)
    return "sem", "", 0.0