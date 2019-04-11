# doc-enhancer
A python script to fix, format, and expand your document. It's main feature is utilizing OpenAI to add content to your paragraphs and setting the format as MLA.

My script Utilizes the open source GPT-2 from OpenAI modified to use Pytorch to use machine learning to create sentances and paragraphs. 

https://github.com/graykode/gpt-2-Pytorch

https://github.com/openai/gpt-2

## Installing and Testing GPT-2
**To get started:**

* Install pytorch from here: https://pytorch.org/

* Install dependencies:
`pip3 install tqdm regex==2017.4.5`

* Clone gpt-2:
`git clone https://github.com/graykode/gpt-2-Pytorch && mv gpt-2-Pytorch gpt2Pytorch && cd gpt2Pytorch`

* Download the 117M model:
`curl --output gpt2-pytorch_model.bin https://s3.amazonaws.com/models.huggingface.co/bert/gpt2-pytorch_model.bin`

* Test it out:
`python main.py --text "It was a bright cold day in April, and the clocks were striking thirteen. Winston Smith, his chin nuzzled into his breast in an effort to escape the vile wind, slipped quickly through the glass doors of Victory Mansions, though not quickly enough to prevent a swirl of gritty dust from entering along with him."`

## Enhancing your document

Prerequisites: `pip install python-docx==0.8.10`

* Save your document in the working directory as "test.docx"

* Run `python3 main2.py`

* The results will be in "mla.docx
