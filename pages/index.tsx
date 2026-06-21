import dynamic from "next/dynamic";

const App = dynamic(() => import("../remixed-8846c3ed"), { ssr: false });

export default function Home() {
  return <App />;
}
