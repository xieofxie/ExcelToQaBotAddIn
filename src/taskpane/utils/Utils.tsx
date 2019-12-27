export async function getConfig(context: Excel.RequestContext) {
    const configSheet = context.workbook.worksheets.getFirst();
    const configRange = configSheet.getUsedRange();
    configRange.load("values");

    await context.sync();

    let result = new Map();
    for (let i = 0;i < configRange.values.length;++i) {
      let element = configRange.values[i];
      if (element.length < 2) continue;
      result.set(String(element[0]).toLowerCase(), element[1]);
    }
    return result;
}
