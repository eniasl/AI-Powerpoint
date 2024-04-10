import os
import openai
import config
openai.api_key = config.OPEN_AI_KEY


def produceQuery(mainpoints, topic):
    response = openai.Completion.create(
      model="text-davinci-003",
      prompt=f"make a 15 minute detailed presentation about {topic} with {mainpoints} as the main points and add a '~' before every sub-point. mark every main point with a hashtag.",
      temperature=0.7,
      max_tokens=256,
      top_p=1,
      frequency_penalty=0,
      presence_penalty=0 )
    if "choices" in response:
        if len(response["choices"])>0:
           print("working")
           answer=response["choices"][0]["text"]
           return answer
        else:
            return "not working"

