import anthropic

client = anthropic.Anthropic(
    api_key="s"
)

text = input("Digite uma mensagem para o Claude: ")

message = client.messages.create(
  model="claude-sonnet-4-6",
  max_tokens=1024,
  messages=[{
    "role": "user",
    "content": text
  }]
)
print(message.content[0].text)