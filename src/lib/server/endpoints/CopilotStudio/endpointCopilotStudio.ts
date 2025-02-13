import type { TextGenerationStreamOutput } from "@huggingface/inference";
import type { Endpoint } from "../endpoints";
import { z } from "zod";

// Schema for Direct Line API configuration
export const endpointCopilotParametersSchema = z.object({
	weight: z.number().int().positive().default(1),
	type: z.literal("copilot"),
	directLineSecret: z.string().default("<direct line secret>"), // Direct Line Secret (Used to generate a token)
	url: z.string().url().default("https://directline.botframework.com/v3/directline"),
});

export function endpointCopilot(input: z.input<typeof endpointCopilotParametersSchema>): Endpoint {
	const { url, directLineSecret } = endpointCopilotParametersSchema.parse(input);

	return async ({ messages, continueMessage, preprompt }) => {
		// Step 1: Generate Direct Line Token
		const tokenResponse = await fetch(`${url}/tokens/generate`, {
			method: "POST",
			headers: {
				"Authorization": `Bearer ${directLineSecret}`,
				"Content-Type": "application/json",
			},
		});

		if (!tokenResponse.ok) {
			throw new Error("Failed to generate Direct Line token.");
		}

		const { token } = await tokenResponse.json();

		// Step 2: Start a new conversation
		const conversationResponse = await fetch(`${url}/conversations`, {
			method: "POST",
			headers: {
				"Authorization": `Bearer ${token}`,
				"Content-Type": "application/json",
			},
		});

		if (!conversationResponse.ok) {
			throw new Error("Failed to start a conversation with Copilot Studio.");
		}

		const { conversationId } = await conversationResponse.json();

		// Step 3: Construct the message
		const prompt = messages.map(m => m.content).join("\n");

		// Step 4: Send user message
		await fetch(`${url}/conversations/${conversationId}/activities`, {
			method: "POST",
			headers: {
				"Authorization": `Bearer ${token}`,
				"Content-Type": "application/json",
			},
			body: JSON.stringify({
				type: "message",
				from: { id: "user" },
				text: prompt,
			}),
		});

		// Step 5: Polling for responses
		return (async function* () {
			let stop = false;
			let tokenId = 0;
			let generatedText = "";

			while (!stop) {
				const response = await fetch(`${url}/conversations/${conversationId}/activities`, {
					method: "GET",
					headers: {
						"Authorization": `Bearer ${token}`,
						"Content-Type": "application/json",
					},
				});

				if (!response.ok) {
					throw new Error("Failed to retrieve messages from Copilot Studio.");
				}

				const data = await response.json();

				// Extract bot responses
				const botMessages = data.activities.filter((activity: any) => activity.from.id !== "user");

				for (const botMessage of botMessages) {
					const text = botMessage.text || "";

					generatedText += text;

					yield {
						token: {
							id: tokenId++,
							text: text,
							logprob: 0,
							special: false,
						},
						generated_text: null,
						details: null,
					} satisfies TextGenerationStreamOutput;
				}

				// Stop when conversation is over
				if (botMessages.length > 0) {
					stop = true;
					yield {
						token: {
							id: tokenId++,
							text: generatedText,
							logprob: 0,
							special: true,
						},
						generated_text: generatedText,
						details: null,
					} satisfies TextGenerationStreamOutput;
				}
			}
		})();
	};
}

export default endpointCopilot;
