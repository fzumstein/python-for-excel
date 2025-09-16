FROM ghcr.io/astral-sh/uv:python3.13-trixie

COPY .python-version .
COPY pyproject.toml .
COPY uv.lock .

RUN uv sync --locked

# create user with a home directory
ARG NB_USER
ARG NB_UID
ENV USER ${NB_USER}
ENV HOME /home/${NB_USER}

RUN adduser --disabled-password \
    --gecos "Default user" \
    --uid ${NB_UID} \
    ${NB_USER}
WORKDIR ${HOME}
USER ${USER}