import json

from PIL import Image, ImageTk


def selecionar_regiao(path):
    import tkinter as tk
    from tkinter import ttk

    with Image.open(path) as original:
        imagem = original.convert("RGB")

    janela = tk.Tk()
    janela.title("Calibrar região do AR")
    imagem.thumbnail((janela.winfo_screenwidth() - 60, janela.winfo_screenheight() - 170))
    foto = ImageTk.PhotoImage(imagem)
    largura, altura = imagem.size
    painel = tk.Canvas(janela, width=largura, height=altura, highlightthickness=0, cursor="crosshair")
    painel.pack()
    painel.create_image(0, 0, anchor="nw", image=foto)
    retangulo = painel.create_rectangle(0, 0, 0, 0, outline="#ff2e4f", width=2)
    valor = tk.StringVar()
    ttk.Label(janela, text="Arraste o mouse do início ao fim da região e copie o valor gerado.").pack(pady=4)
    ttk.Entry(janela, textvariable=valor, state="readonly", width=55).pack(pady=4)

    def copiar():
        if valor.get():
            janela.clipboard_clear()
            janela.clipboard_append(valor.get())

    ttk.Button(janela, text="Copiar região", command=copiar).pack(pady=4)
    inicio = None

    def coordenadas(event):
        return min(max(event.x, 0), largura), min(max(event.y, 0), altura)

    def comecar(event):
        nonlocal inicio
        inicio = coordenadas(event)
        painel.coords(retangulo, *inicio, *inicio)

    def desenhar(event):
        if inicio is not None:
            painel.coords(retangulo, *inicio, *coordenadas(event))

    def concluir(event):
        desenhar(event)
        if inicio is None:
            return
        fim = coordenadas(event)
        esq, dire = sorted((inicio[0], fim[0]))
        topo, base = sorted((inicio[1], fim[1]))
        if esq == dire or topo == base:
            return
        resultado = [esq / largura, topo / altura, dire / largura, base / altura]
        valor.set(json.dumps([round(ponto, 5) for ponto in resultado]))

    painel.bind("<ButtonPress-1>", comecar)
    painel.bind("<B1-Motion>", desenhar)
    painel.bind("<ButtonRelease-1>", concluir)
    janela.mainloop()
