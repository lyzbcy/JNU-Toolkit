(async (base) => {
  const normalize = text => text.replace(/\s+/g, " ").trim();
  const getDocument = async url => {
    let lastError;
    for (let attempt = 0; attempt < 2; attempt += 1) {
      try {
        const response = await fetch(url, {
          cache: "no-store",
          signal: AbortSignal.timeout(11000),
        });
        const html = await response.text();
        if (!response.ok || html.trim().length < 30) {
          throw new Error("HTTP " + response.status + "，页面内容为空或加载失败");
        }
        return { url: response.url, doc: new DOMParser().parseFromString(html, "text/html") };
      } catch (error) {
        lastError = error;
      }
    }
    throw lastError;
  };

  const courses = [...document.querySelectorAll('.courselist a[href*="courseID="]')].map(link => ({
    id: new URL(link.getAttribute("href"), base).searchParams.get("courseID"),
    name: normalize(link.textContent),
  }));
  if (!courses.length) throw new Error("找不到已登录的课程列表");

  const readAssignment = async link => {
    const { doc } = await getDocument(link.url);
    const content = doc.querySelector(".col-10");
    const heading = content?.querySelector("h4");
    const dateLine = [...(content?.querySelectorAll("p") ?? [])]
      .find(p => normalize(p.textContent).startsWith("作业时间"));
    const dates = [...(dateLine?.querySelectorAll("b") ?? [])]
      .map(date => normalize(date.textContent));
    if (dates.length !== 2 || normalize(heading?.textContent ?? "") !== link.name) {
      throw new Error("作业标题或开闭时间与列表不匹配");
    }
    const questions = [...content.querySelectorAll("table tbody tr")]
      .filter(row => row.cells.length >= 4)
      .map((row) => {
        const [, name, points, statusCell] = row.cells;
        const status = statusCell.cloneNode(true);
        status.querySelectorAll("a").forEach(link => link.remove());
        status.querySelectorAll("p, div, br, pre, li").forEach(
          node => node.after(status.ownerDocument.createTextNode(" ")),
        );
        const full = normalize(status.textContent);
        const summary = full.split(/文件重命名为|RANK:|STDOUT:|STDERR:/i)[0].trim();
        const exitCode = full.match(/RETURN CODE:\s*(-?\d+)/i)?.[1];
        return {
          name: normalize(name.textContent),
          points: normalize(points.textContent),
          status: exitCode === undefined ? summary : summary + " RETURN CODE: " + exitCode,
        };
      });
    if (!questions.length) throw new Error("作业的题目状态表为空");
    return { ...link, opens: dates[0], deadline: dates[1], questions };
  };

  const state = globalThis.cgHomeworkScan = { done: false, running: null, courses: [] };
  // The course switch updates the server-side session, so read courses in order.
  for (const course of courses) {
    const result = { ...course, assignments: [] };
    state.running = course.name;
    try {
      const selection = await getDocument(new URL("courselist.jsp?courseID=" + course.id, base));
      if (!new URL(selection.url).pathname.endsWith("/main.jsp")) {
        throw new Error("切换课程后未进入课程主页");
      }
      const page = await getDocument(new URL("includes/redirect.jsp?tab=-2", base));
      const current = [...page.doc.querySelectorAll("span")]
        .find(node => normalize(node.textContent) === "当前作业")?.nextElementSibling;
      if (!current?.classList.contains("list-group")) throw new Error("当前作业菜单未加载");
      const courseIds = [...current.parentElement.querySelectorAll('a[href*="assignID"]')]
        .map(link => new URL(link.getAttribute("href"), page.url).searchParams.get("courseID"))
        .filter(Boolean);
      if (courseIds.some(id => id !== course.id)) throw new Error("作业菜单属于其他课程");

      for (const link of current.querySelectorAll('a[href*="assignID"]')) {
        const assignment = {
          name: normalize(link.textContent),
          url: new URL(link.getAttribute("href"), page.url).href,
        };
        try {
          result.assignments.push(await readAssignment(assignment));
        } catch (error) {
          result.assignments.push({ ...assignment, error: String(error) });
        }
      }
    } catch (error) {
      result.error = String(error);
    }
    state.courses.push(result);
  }
  state.running = null;
  state.done = true;
})
