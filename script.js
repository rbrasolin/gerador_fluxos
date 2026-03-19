function escolherParesCandidatos(origem, destino, rotulo = "") {
  if (origem.isDecision && rotulo === "Sim") {
    return [{ startSide: "right", endSide: "left" }];
  }

  if (origem.isDecision && rotulo === "Não") {
    return [{ startSide: "bottom", endSide: "top" }];
  }

  const dx = destino.gridCol - origem.gridCol;
  const dy = destino.gridRow - origem.gridRow;

  // mesma coluna
  if (dx === 0 && dy > 0) {
    return [
      { startSide: "bottom", endSide: "top" }
    ];
  }

  if (dx === 0 && dy < 0) {
    return [
      { startSide: "top", endSide: "bottom" }
    ];
  }

  // mesma linha
  if (dy === 0 && dx > 0) {
    return [
      { startSide: "right", endSide: "left" }
    ];
  }

  if (dy === 0 && dx < 0) {
    return [
      { startSide: "left", endSide: "right" }
    ];
  }

  // diagonal para a direita -> SEMPRE sai pela direita
  if (dx > 0 && dy > 0) {
    return [
      { startSide: "right", endSide: "left" },
      { startSide: "right", endSide: "top" },
      { startSide: "bottom", endSide: "left" },
      { startSide: "bottom", endSide: "top" }
    ];
  }

  if (dx > 0 && dy < 0) {
    return [
      { startSide: "right", endSide: "left" },
      { startSide: "right", endSide: "bottom" },
      { startSide: "top", endSide: "left" },
      { startSide: "top", endSide: "bottom" }
    ];
  }

  // diagonal para a esquerda -> SEMPRE sai pela esquerda
  if (dx < 0 && dy > 0) {
    return [
      { startSide: "left", endSide: "right" },
      { startSide: "left", endSide: "top" },
      { startSide: "bottom", endSide: "right" },
      { startSide: "bottom", endSide: "top" }
    ];
  }

  if (dx < 0 && dy < 0) {
    return [
      { startSide: "left", endSide: "right" },
      { startSide: "left", endSide: "bottom" },
      { startSide: "top", endSide: "right" },
      { startSide: "top", endSide: "bottom" }
    ];
  }

  return [{ startSide: "right", endSide: "left" }];
}