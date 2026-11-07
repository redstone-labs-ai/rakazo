/** The name bots call the product in prompts and in the chat messages the server writes. */
export const DEFAULT_PRODUCT_NAME = "Rakazo";

export const MAX_PRODUCT_NAME_LENGTH = 40;

/** RAKAZO_PRODUCT_NAME, trimmed. Unset or blank keeps the default name. */
export function productNameFromEnv(source: { RAKAZO_PRODUCT_NAME?: string }): string {
  const name = source.RAKAZO_PRODUCT_NAME?.trim() || DEFAULT_PRODUCT_NAME;
  if (name.length > MAX_PRODUCT_NAME_LENGTH) {
    throw new Error(`RAKAZO_PRODUCT_NAME must be at most ${MAX_PRODUCT_NAME_LENGTH} characters`);
  }
  return name;
}
