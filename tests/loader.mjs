export async function resolve(specifier, context, nextResolve) {
  const clean = specifier.replace(/\?v=\d+/, '');
  return nextResolve(clean, context);
}
